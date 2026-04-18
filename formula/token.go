package formula

import "fmt"

// TokenType classifies a lexer token.
type TokenType byte

const (
	TokEOF        TokenType = iota
	TokNumber               // 123, 1.5, 1.5E10
	TokString               // "hello"
	TokBool                 // TRUE, FALSE
	TokError                // #N/A, #DIV/0!, etc.
	TokCellRef              // A1, $A$1, Sheet1!A1, 'Sheet Name'!A1
	TokFunc                 // SUM(, IF(  — name includes the opening paren
	TokOp                   // +, -, *, /, ^, &, =, <>, <, >, <=, >=
	TokLParen               // (
	TokRParen               // )
	TokComma                // ,
	TokSemicolon            // ; (row separator in array literals)
	TokColon                // : (range operator, e.g. A1:B5)
	TokPercent              // %
	TokArrayOpen            // {
	TokArrayClose           // }
	TokBang                 // ! (sheet separator, when not part of a cell ref)
	TokIntersect            // whitespace intersection operator between references (Excel)
)

// Token is a single token produced by the lexer.
type Token struct {
	Type  TokenType
	Value string // raw text of the token
	Pos   int    // byte offset in the source formula
}

func (t Token) String() string {
	return fmt.Sprintf("{%s %q @%d}", t.Type, t.Value, t.Pos)
}

var tokenTypeNames = [...]string{
	TokEOF:        "EOF",
	TokNumber:     "Number",
	TokString:     "String",
	TokBool:       "Bool",
	TokError:      "Error",
	TokCellRef:    "CellRef",
	TokFunc:       "Func",
	TokOp:         "Op",
	TokLParen:     "LParen",
	TokRParen:     "RParen",
	TokComma:      "Comma",
	TokSemicolon:  "Semicolon",
	TokColon:      "Colon",
	TokPercent:    "Percent",
	TokArrayOpen:  "ArrayOpen",
	TokArrayClose: "ArrayClose",
	TokBang:       "Bang",
	TokIntersect:  "Intersect",
}

func (t TokenType) String() string {
	if int(t) < len(tokenTypeNames) {
		return tokenTypeNames[t]
	}
	return fmt.Sprintf("TokenType(%d)", t)
}
