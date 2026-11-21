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
	"^":  {10, 11}, // left-associative (matches expected behavior)
}

const (
	colonLeftBP      = 14
	colonRightBP     = 15
	intersectLeftBP  = 13 // space intersection binds slightly weaker than colon
	intersectRightBP = 14
	prefixRBP        = 11 // unary - and + bind tighter than ^ (convention: -2^2 = 4)
	postfixBP        = 12 // postfix % — above binary ops, below range colon

	maxRow = maxRows // maximum row number
	maxCol = maxCols // maximum column number (XFD)
)

// parseExpression is the core Pratt parsing loop.
func (p *Parser) parseExpression(minBP int) (Node, error) {
	left, err := p.parseNud()
	if err != nil {
		return nil, err
	}

	for {
		tok := p.peek()

		if tok.Type == TokPercent {
			// Postfix % binds at postfixBP. Gated by minBP so that the RHS
			// of a colon range (parsed at colonRightBP=15) leaves % to the
			// surrounding range expression, e.g. A1:A4% parses as (A1:A4)%.
			if postfixBP < minBP {
				break
			}
			p.advance()
			left = &PostfixExpr{Op: "%", Operand: left}
			continue
		}

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
			continue
		}

		if tok.Type == TokIntersect {
			if intersectLeftBP < minBP {
				break
			}
			p.advance()
			right, err := p.parseExpression(intersectRightBP)
			if err != nil {
				return nil, err
			}
			left = &IntersectRef{Left: left, Right: right}
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
				} else if fromOK && toIsNum && fromRef.Col == 0 {
					// Sheet-qualified row-only range like Sheet1!2:3. The
					// left side is a CellRef with Col==0 (row-only); the
					// right side parsed as a bare number and needs to be
					// lifted into a matching CellRef that inherits the
					// sheet qualifier.
					toRow := int(toNum.Value)
					if toRow < 1 || float64(toRow) != toNum.Value {
						return nil, fmt.Errorf("invalid row range endpoint %s", toNum.Raw)
					}
					toRef = &CellRef{Col: 0, Row: toRow}
					toOK = true
				} else {
					// If one side is already an error (e.g. from a cross-sheet
					// range that produced #VALUE!), propagate it instead of
					// returning a parse error. This allows COUNT(#VALUE!) to
					// return 0 instead of failing the entire formula.
					if _, leftIsErr := left.(*ErrorLit); leftIsErr {
						continue
					}
					// Allow ref-returning expressions on either side
					// (e.g. A1:INDEX(A:A,n) or OFFSET(A1,1,0):B5). Both
					// endpoints must be reference-producing — cells,
					// ranges, or functions like INDEX/OFFSET/INDIRECT/
					// CHOOSE/IF/IFS/SWITCH/ANCHORARRAY.
					if isRefProducingNode(left) && isRefProducingNode(right) {
						left = &DynamicRangeRef{From: left, To: right}
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
			// (e.g. S1:S3!A1 or Sheet1!A1:Sheet2!B5). Return #VALUE!.
			// Note: Sheet1!A1:B5 is valid — B5 has no sheet and inherits Sheet1.
			if toRef.Sheet != "" && toRef.Sheet != fromRef.Sheet {
				left = &ErrorLit{Code: ErrVALUE}
				continue
			}

			// Expand column-only references (Row==0) into full-column ranges.
			// F:F becomes F1:F1048576.
			if fromRef.Row == 0 {
				fromRef.Row = 1
			}
			if toRef.Row == 0 {
				toRef.Row = maxRow
			}
			// Expand row-only references (Col==0) into full-row ranges.
			// 5:6 becomes A5:XFD6.
			if fromRef.Col == 0 {
				fromRef.Col = 1
			}
			if toRef.Col == 0 {
				toRef.Col = maxCol
			}
			left = &RangeRef{From: fromRef, To: toRef}
			continue
		}

		break
	}

	return left, nil
}

// isRefProducingNode reports whether a node yields a cell or range reference,
// either statically (CellRef/RangeRef/IntersectRef/UnionRef) or via a
// reference-returning function call (INDEX/OFFSET/INDIRECT/CHOOSE/IF/IFS/
// SWITCH/ANCHORARRAY). Used to determine whether `LEFT:RIGHT` should be
// parsed as a DynamicRangeRef when the static cell-ref form doesn't apply.
func isRefProducingNode(n Node) bool {
	switch v := n.(type) {
	case *CellRef, *RangeRef, *IntersectRef, *UnionRef, *DynamicRangeRef:
		return true
	case *FuncCall:
		switch strings.ToUpper(v.Name) {
		case "INDEX", "OFFSET", "INDIRECT", "CHOOSE", "IF", "IFS", "SWITCH", "ANCHORARRAY":
			return true
		}
	}
	return false
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
		// A parenthesized list of references ( A1:A2, C1:C2, ... ) is a union
		// reference when every element is reference-like.
		if p.peek().Type == TokComma && isUnionReferenceNode(expr) {
			areas := []Node{expr}
			for p.peek().Type == TokComma {
				p.advance()
				area, err := p.parseExpression(0)
				if err != nil {
					return nil, err
				}
				if !isUnionReferenceNode(area) {
					return nil, fmt.Errorf("union references must contain only direct cell or range references")
				}
				areas = append(areas, area)
			}
			if _, err := p.expect(TokRParen); err != nil {
				return nil, fmt.Errorf("expected ')' to close reference list")
			}
			return &UnionRef{Areas: areas}, nil
		}
		if _, err := p.expect(TokRParen); err != nil {
			return nil, fmt.Errorf("unmatched '(' at position %d", tok.Pos)
		}
		return expr, nil

	case TokAt:
		p.advance()
		operand, err := p.parseExpression(prefixRBP)
		if err != nil {
			return nil, err
		}
		return &UnaryExpr{Op: "@", Operand: operand}, nil

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
	upperName := strings.ToUpper(name)

	// Zero-arg function: immediately followed by ).
	if p.peek().Type == TokRParen {
		p.advance()
		if isMapFuncName(name) {
			return desugarMAP(nil)
		}
		return &FuncCall{Name: name}, nil
	}

	var args []Node

	// Handle first argument: may be empty (e.g. SUM(,1))
	if p.peek().Type == TokComma {
		args = append(args, &EmptyArg{})
	} else {
		var (
			arg Node
			err error
		)
		if upperName == "AREAS" {
			arg, err = p.parseAREASArg()
		} else {
			arg, err = p.parseExpression(0)
		}
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

	if isLetFuncName(name) {
		return desugarLET(args)
	}

	if isMapFuncName(name) {
		return desugarMAP(args)
	}

	if isReduceFuncName(name) {
		return desugarREDUCE(args)
	}

	if isScanFuncName(name) {
		return desugarSCAN(args)
	}

	if isByRowFuncName(name) {
		return desugarBYROW(args)
	}

	if isByColFuncName(name) {
		return desugarBYCOL(args)
	}

	if isMakeArrayFuncName(name) {
		return desugarMAKEARRAY(args)
	}

	return call, nil
}

func (p *Parser) parseAREASArg() (Node, error) {
	if p.peek().Type != TokLParen {
		return p.parseExpression(0)
	}
	start := p.pos
	arg, ok, err := p.tryParseAREASUnion()
	if err != nil {
		return nil, err
	}
	if ok {
		return arg, nil
	}
	p.pos = start
	return p.parseExpression(0)
}

func (p *Parser) tryParseAREASUnion() (Node, bool, error) {
	if p.peek().Type != TokLParen {
		return nil, false, nil
	}
	p.advance()

	first, err := p.parseExpression(0)
	if err != nil {
		return nil, false, err
	}
	if p.peek().Type != TokComma {
		if _, err := p.expect(TokRParen); err != nil {
			return nil, false, err
		}
		return first, true, nil
	}
	if !isAREASReferenceNode(first) {
		return nil, false, fmt.Errorf("AREAS multi-area references must contain only direct cell or range references")
	}

	areas := []Node{first}
	for p.peek().Type == TokComma {
		p.advance()
		area, err := p.parseExpression(0)
		if err != nil {
			return nil, false, err
		}
		if !isAREASReferenceNode(area) {
			return nil, false, fmt.Errorf("AREAS multi-area references must contain only direct cell or range references")
		}
		areas = append(areas, area)
	}
	if _, err := p.expect(TokRParen); err != nil {
		return nil, false, fmt.Errorf("expected ')' to close AREAS reference list")
	}
	return &UnionRef{Areas: areas}, true, nil
}

func isAREASReferenceNode(n Node) bool {
	switch n.(type) {
	case *CellRef, *RangeRef:
		return true
	default:
		return false
	}
}

// isUnionReferenceNode reports whether a node can appear inside a
// parenthesized union reference list. Matches AREAS's rules plus allows
// nested intersections and unions.
func isUnionReferenceNode(n Node) bool {
	switch n.(type) {
	case *CellRef, *RangeRef, *IntersectRef, *UnionRef:
		return true
	}
	return false
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

func isLetFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "LET" || upper == "_XLFN.LET"
}

func desugarLET(args []Node) (Node, error) {
	if len(args) < 3 || len(args)%2 == 0 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	// Process name/value pairs sequentially.
	for i := 0; i < len(args)-1; i += 2 {
		nameNode := args[i]
		valueNode := args[i+1]

		name, ok := lambdaParamName(nameNode)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}

		// Substitute this name with its value in all subsequent args.
		subst := map[string]Node{name: valueNode}
		for j := i + 2; j < len(args); j++ {
			args[j] = substituteLambdaNames(args[j], subst)
		}
	}

	// Return the final (fully substituted) calculation.
	return args[len(args)-1], nil
}

func isMapFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "MAP" || upper == "_XLFN.MAP"
}

func desugarMAP(args []Node) (Node, error) {
	if len(args) < 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	// Last arg must be a LAMBDA FuncCall
	lastArg := args[len(args)-1]
	lambdaCall, ok := lastArg.(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	// LAMBDA must have at least a body (1 arg) and params matching array count
	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) == 0 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	numArrays := len(args) - 1
	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Number of params must match number of arrays
	if len(params) != numArrays {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	// Extract parameter names and build substitution map
	paramNames := make([]string, len(params))
	subst := make(map[string]Node, len(params))
	for i, param := range params {
		name, ok := lambdaParamName(param)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}
		paramNames[i] = name
		subst[name] = &ParamRef{Slot: i, Name: name}
	}

	// Replace parameter references in the body with ParamRef nodes
	transformedBody := substituteLambdaNames(body, subst)

	return &MapExpr{
		Arrays:     args[:numArrays],
		ParamNames: paramNames,
		Body:       transformedBody,
	}, nil
}

func isReduceFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "REDUCE" || upper == "_XLFN.REDUCE"
}

func desugarREDUCE(args []Node) (Node, error) {
	// REDUCE(initial_value, array, lambda)
	if len(args) != 3 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	initialValue := args[0]
	arrayExpr := args[1]

	// Last arg must be LAMBDA
	lambdaCall, ok := args[2].(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) < 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Must have exactly 2 params: accumulator and value
	if len(params) != 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	paramNames := make([]string, 2)
	subst := make(map[string]Node, 2)
	for i, param := range params {
		name, ok := lambdaParamName(param)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}
		paramNames[i] = name
		subst[name] = &ParamRef{Slot: i, Name: name}
	}

	transformedBody := substituteLambdaNames(body, subst)

	return &ReduceExpr{
		InitialValue: initialValue,
		Array:        arrayExpr,
		ParamNames:   paramNames,
		Body:         transformedBody,
	}, nil
}

func isScanFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "SCAN" || upper == "_XLFN.SCAN"
}

func desugarSCAN(args []Node) (Node, error) {
	// SCAN(initial_value, array, lambda)
	if len(args) != 3 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	initialValue := args[0]
	arrayExpr := args[1]

	// Last arg must be LAMBDA
	lambdaCall, ok := args[2].(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) < 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Must have exactly 2 params: accumulator and value
	if len(params) != 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	paramNames := make([]string, 2)
	subst := make(map[string]Node, 2)
	for i, param := range params {
		name, ok := lambdaParamName(param)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}
		paramNames[i] = name
		subst[name] = &ParamRef{Slot: i, Name: name}
	}

	transformedBody := substituteLambdaNames(body, subst)

	return &ScanExpr{
		InitialValue: initialValue,
		Array:        arrayExpr,
		ParamNames:   paramNames,
		Body:         transformedBody,
	}, nil
}

func isByRowFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "BYROW" || upper == "_XLFN.BYROW"
}

func desugarBYROW(args []Node) (Node, error) {
	// BYROW(array, lambda)
	if len(args) != 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	arrayExpr := args[0]

	// Last arg must be LAMBDA
	lambdaCall, ok := args[1].(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) < 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Must have exactly 1 param
	if len(params) != 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	paramName, ok := lambdaParamName(params[0])
	if !ok {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	subst := map[string]Node{paramName: &ParamRef{Slot: 0, Name: paramName}}
	transformedBody := substituteLambdaNames(body, subst)

	return &ByRowExpr{
		Array:      arrayExpr,
		ParamNames: []string{paramName},
		Body:       transformedBody,
	}, nil
}

func isByColFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "BYCOL" || upper == "_XLFN.BYCOL"
}

func desugarBYCOL(args []Node) (Node, error) {
	// BYCOL(array, lambda)
	if len(args) != 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	arrayExpr := args[0]

	// Last arg must be LAMBDA
	lambdaCall, ok := args[1].(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) < 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Must have exactly 1 param
	if len(params) != 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	paramName, ok := lambdaParamName(params[0])
	if !ok {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	subst := map[string]Node{paramName: &ParamRef{Slot: 0, Name: paramName}}
	transformedBody := substituteLambdaNames(body, subst)

	return &ByColExpr{
		Array:      arrayExpr,
		ParamNames: []string{paramName},
		Body:       transformedBody,
	}, nil
}

func isMakeArrayFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "MAKEARRAY" || upper == "_XLFN.MAKEARRAY"
}

func desugarMAKEARRAY(args []Node) (Node, error) {
	if len(args) != 3 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	rowsExpr := args[0]
	colsExpr := args[1]

	lambdaCall, ok := args[2].(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	lambdaArgs := lambdaCall.Args
	if len(lambdaArgs) < 1 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]

	// Must have exactly 2 params
	if len(params) != 2 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	paramNames := make([]string, 2)
	subst := make(map[string]Node, 2)
	for i, param := range params {
		name, ok := lambdaParamName(param)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}
		paramNames[i] = name
		subst[name] = &ParamRef{Slot: i, Name: name}
	}

	transformedBody := substituteLambdaNames(body, subst)

	return &MakeArrayExpr{
		Rows:       rowsExpr,
		Cols:       colsExpr,
		ParamNames: paramNames,
		Body:       transformedBody,
	}, nil
}

// inlineBoundLambdaCall checks whether funcName refers to a name in subst that
// is bound to a LAMBDA. If so, it desugars the invocation into the lambda's
// body with callArgs substituted for its parameters. Returns (node, true) on a
// successful inline.
func inlineBoundLambdaCall(funcName string, callArgs []Node, subst map[string]Node) (Node, bool) {
	key := strings.ToUpper(funcName)
	key = strings.TrimPrefix(key, "_XLFN._XLWS.")
	key = strings.TrimPrefix(key, "_XLFN.")
	key = strings.TrimPrefix(key, "_XLPM.")
	repl, ok := subst[key]
	if !ok {
		return nil, false
	}
	lambdaCall, ok := repl.(*FuncCall)
	if !ok || !isLambdaFuncName(lambdaCall.Name) {
		return nil, false
	}
	lambdaArgs := make([]Node, len(lambdaCall.Args))
	for i, a := range lambdaCall.Args {
		lambdaArgs[i] = cloneNode(a)
	}
	node, err := desugarLambdaInvocation(lambdaArgs, callArgs)
	if err != nil {
		return nil, false
	}
	return node, true
}

func lambdaParamName(n Node) (string, bool) {
	ref, ok := n.(*CellRef)
	if !ok || ref.Sheet != "" || ref.SheetEnd != "" || ref.AbsCol || ref.AbsRow || ref.DotNotation {
		return "", false
	}
	if ref.Name != "" {
		return strings.ToUpper(ref.Name), true
	}
	if ref.Row != 0 {
		return "", false
	}
	return strings.ToUpper(ColNumberToLetters(ref.Col)), true
}

func substituteLambdaNames(n Node, subst map[string]Node) Node {
	switch v := n.(type) {
	case *CellRef:
		if v.Sheet == "" && v.SheetEnd == "" && !v.AbsCol && !v.AbsRow && !v.DotNotation {
			if v.Name != "" {
				if repl, ok := subst[strings.ToUpper(v.Name)]; ok {
					return cloneNode(repl)
				}
			} else if v.Row == 0 {
				if repl, ok := subst[strings.ToUpper(ColNumberToLetters(v.Col))]; ok {
					return cloneNode(repl)
				}
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
		// If the function name itself refers to a LET-bound LAMBDA, inline
		// the invocation (e.g. LET(sq, LAMBDA(n,n*n), sq(A1)) → (n*n)[n:=A1]).
		if inlined, ok := inlineBoundLambdaCall(v.Name, args, subst); ok {
			return inlined
		}
		return &FuncCall{Name: v.Name, Args: args}
	case *UnionRef:
		areas := make([]Node, len(v.Areas))
		for i, area := range v.Areas {
			areas[i] = substituteLambdaNames(area, subst)
		}
		return &UnionRef{Areas: areas}
	case *ArrayLit:
		rows := make([][]Node, len(v.Rows))
		for i, row := range v.Rows {
			rows[i] = make([]Node, len(row))
			for j, elem := range row {
				rows[i][j] = substituteLambdaNames(elem, subst)
			}
		}
		return &ArrayLit{Rows: rows}
	case *ParamRef:
		// ParamRef is already a resolved parameter reference; return as-is.
		return &ParamRef{Slot: v.Slot, Name: v.Name}
	case *MapExpr:
		arrays := make([]Node, len(v.Arrays))
		for i, arr := range v.Arrays {
			arrays[i] = substituteLambdaNames(arr, subst)
		}
		return &MapExpr{
			Arrays:     arrays,
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       substituteLambdaNames(v.Body, subst),
		}
	case *ReduceExpr:
		return &ReduceExpr{
			InitialValue: substituteLambdaNames(v.InitialValue, subst),
			Array:        substituteLambdaNames(v.Array, subst),
			ParamNames:   append([]string(nil), v.ParamNames...),
			Body:         substituteLambdaNames(v.Body, subst),
		}
	case *ScanExpr:
		return &ScanExpr{
			InitialValue: substituteLambdaNames(v.InitialValue, subst),
			Array:        substituteLambdaNames(v.Array, subst),
			ParamNames:   append([]string(nil), v.ParamNames...),
			Body:         substituteLambdaNames(v.Body, subst),
		}
	case *ByRowExpr:
		return &ByRowExpr{
			Array:      substituteLambdaNames(v.Array, subst),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       substituteLambdaNames(v.Body, subst),
		}
	case *ByColExpr:
		return &ByColExpr{
			Array:      substituteLambdaNames(v.Array, subst),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       substituteLambdaNames(v.Body, subst),
		}
	case *MakeArrayExpr:
		return &MakeArrayExpr{
			Rows:       substituteLambdaNames(v.Rows, subst),
			Cols:       substituteLambdaNames(v.Cols, subst),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       substituteLambdaNames(v.Body, subst),
		}
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
	case *UnionRef:
		areas := make([]Node, len(v.Areas))
		for i, area := range v.Areas {
			areas[i] = cloneNode(area)
		}
		return &UnionRef{Areas: areas}
	case *ArrayLit:
		rows := make([][]Node, len(v.Rows))
		for i, row := range v.Rows {
			rows[i] = make([]Node, len(row))
			for j, elem := range row {
				rows[i][j] = cloneNode(elem)
			}
		}
		return &ArrayLit{Rows: rows}
	case *ParamRef:
		return &ParamRef{Slot: v.Slot, Name: v.Name}
	case *MapExpr:
		arrays := make([]Node, len(v.Arrays))
		for i, arr := range v.Arrays {
			arrays[i] = cloneNode(arr)
		}
		return &MapExpr{
			Arrays:     arrays,
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       cloneNode(v.Body),
		}
	case *ReduceExpr:
		return &ReduceExpr{
			InitialValue: cloneNode(v.InitialValue),
			Array:        cloneNode(v.Array),
			ParamNames:   append([]string(nil), v.ParamNames...),
			Body:         cloneNode(v.Body),
		}
	case *ScanExpr:
		return &ScanExpr{
			InitialValue: cloneNode(v.InitialValue),
			Array:        cloneNode(v.Array),
			ParamNames:   append([]string(nil), v.ParamNames...),
			Body:         cloneNode(v.Body),
		}
	case *ByRowExpr:
		return &ByRowExpr{
			Array:      cloneNode(v.Array),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       cloneNode(v.Body),
		}
	case *ByColExpr:
		return &ByColExpr{
			Array:      cloneNode(v.Array),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       cloneNode(v.Body),
		}
	case *MakeArrayExpr:
		return &MakeArrayExpr{
			Rows:       cloneNode(v.Rows),
			Cols:       cloneNode(v.Cols),
			ParamNames: append([]string(nil), v.ParamNames...),
			Body:       cloneNode(v.Body),
		}
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
