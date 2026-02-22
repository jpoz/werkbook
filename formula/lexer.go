package formula

import (
	"fmt"
	"strings"
)

// Lexer tokenizes an Excel formula string.
type Lexer struct {
	src []byte
	pos int // current byte position
}

// NewLexer creates a lexer for the given formula string.
// The formula should NOT include the leading '=' that Excel uses;
// strip it before passing to the lexer.
func NewLexer(formula string) *Lexer {
	return &Lexer{src: []byte(formula)}
}

// Tokenize returns all tokens for the formula, ending with TokEOF.
func Tokenize(formula string) ([]Token, error) {
	l := NewLexer(formula)
	var tokens []Token
	for {
		tok, err := l.Next()
		if err != nil {
			return nil, err
		}
		tokens = append(tokens, tok)
		if tok.Type == TokEOF {
			break
		}
	}
	return tokens, nil
}

// Next returns the next token from the input.
func (l *Lexer) Next() (Token, error) {
	l.skipWhitespace()

	if l.pos >= len(l.src) {
		return Token{Type: TokEOF, Pos: l.pos}, nil
	}

	ch := l.src[l.pos]

	// String literal: "..."
	if ch == '"' {
		return l.lexString()
	}

	// Error literal: #DIV/0!, #N/A, etc.
	if ch == '#' {
		return l.lexError()
	}

	// Array braces.
	if ch == '{' {
		tok := Token{Type: TokArrayOpen, Value: "{", Pos: l.pos}
		l.pos++
		return tok, nil
	}
	if ch == '}' {
		tok := Token{Type: TokArrayClose, Value: "}", Pos: l.pos}
		l.pos++
		return tok, nil
	}

	// Parentheses.
	if ch == '(' {
		tok := Token{Type: TokLParen, Value: "(", Pos: l.pos}
		l.pos++
		return tok, nil
	}
	if ch == ')' {
		tok := Token{Type: TokRParen, Value: ")", Pos: l.pos}
		l.pos++
		return tok, nil
	}

	// Punctuation.
	if ch == ',' {
		tok := Token{Type: TokComma, Value: ",", Pos: l.pos}
		l.pos++
		return tok, nil
	}
	if ch == ';' {
		tok := Token{Type: TokSemicolon, Value: ";", Pos: l.pos}
		l.pos++
		return tok, nil
	}
	if ch == ':' {
		tok := Token{Type: TokColon, Value: ":", Pos: l.pos}
		l.pos++
		return tok, nil
	}
	if ch == '%' {
		tok := Token{Type: TokPercent, Value: "%", Pos: l.pos}
		l.pos++
		return tok, nil
	}

	// Multi-char comparison operators.
	if ch == '<' {
		return l.lexComparisonOp()
	}
	if ch == '>' {
		return l.lexComparisonOp()
	}
	if ch == '=' {
		tok := Token{Type: TokOp, Value: "=", Pos: l.pos}
		l.pos++
		return tok, nil
	}

	// Arithmetic and concat operators.
	if ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^' || ch == '&' {
		tok := Token{Type: TokOp, Value: string(ch), Pos: l.pos}
		l.pos++
		return tok, nil
	}

	// Quoted sheet name: 'Sheet Name'!A1 or 'Sheet Name'!A1:B5
	if ch == '\'' {
		return l.lexQuotedRef()
	}

	// Number literal: digits or decimal point.
	if ch >= '0' && ch <= '9' || (ch == '.' && l.pos+1 < len(l.src) && l.src[l.pos+1] >= '0' && l.src[l.pos+1] <= '9') {
		return l.lexNumber()
	}

	// Identifier: cell ref, function name, bool, or sheet-qualified ref.
	if isIdentStart(ch) {
		return l.lexIdentOrRef()
	}

	// Bang (sheet separator) when standalone.
	if ch == '!' {
		tok := Token{Type: TokBang, Value: "!", Pos: l.pos}
		l.pos++
		return tok, nil
	}

	return Token{}, fmt.Errorf("unexpected character %q at position %d", ch, l.pos)
}

// lexString reads a double-quoted string. Embedded quotes are doubled ("").
func (l *Lexer) lexString() (Token, error) {
	start := l.pos
	l.pos++ // skip opening quote
	var b strings.Builder
	for l.pos < len(l.src) {
		ch := l.src[l.pos]
		if ch == '"' {
			l.pos++
			// Doubled quote => escaped quote character.
			if l.pos < len(l.src) && l.src[l.pos] == '"' {
				b.WriteByte('"')
				l.pos++
				continue
			}
			// End of string.
			return Token{Type: TokString, Value: b.String(), Pos: start}, nil
		}
		b.WriteByte(ch)
		l.pos++
	}
	return Token{}, fmt.Errorf("unterminated string starting at position %d", start)
}

// lexError reads an error literal like #N/A, #DIV/0!, #REF!, etc.
func (l *Lexer) lexError() (Token, error) {
	start := l.pos
	l.pos++ // skip #

	// Collect until we find the terminator (! or ?) or run out of valid chars.
	for l.pos < len(l.src) {
		ch := l.src[l.pos]
		if ch == '!' || ch == '?' {
			l.pos++
			val := string(l.src[start:l.pos])
			if isExcelError(val) {
				return Token{Type: TokError, Value: val, Pos: start}, nil
			}
			return Token{}, fmt.Errorf("unknown error literal %q at position %d", val, start)
		}
		if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch == '/' || ch == '0' {
			l.pos++
			continue
		}
		break
	}
	// #N/A has no trailing ! or ?
	val := string(l.src[start:l.pos])
	if isExcelError(val) {
		return Token{Type: TokError, Value: val, Pos: start}, nil
	}
	return Token{}, fmt.Errorf("unknown error literal %q at position %d", val, start)
}

// lexComparisonOp reads <, >, <=, >=, <>.
func (l *Lexer) lexComparisonOp() (Token, error) {
	start := l.pos
	ch := l.src[l.pos]
	l.pos++
	if l.pos < len(l.src) {
		next := l.src[l.pos]
		if ch == '<' && (next == '=' || next == '>') {
			l.pos++
			return Token{Type: TokOp, Value: string(l.src[start:l.pos]), Pos: start}, nil
		}
		if ch == '>' && next == '=' {
			l.pos++
			return Token{Type: TokOp, Value: string(l.src[start:l.pos]), Pos: start}, nil
		}
	}
	return Token{Type: TokOp, Value: string(ch), Pos: start}, nil
}

// lexNumber reads a numeric literal, including decimals and scientific notation.
func (l *Lexer) lexNumber() (Token, error) {
	start := l.pos
	l.consumeDigits()

	// Decimal part.
	if l.pos < len(l.src) && l.src[l.pos] == '.' {
		l.pos++
		l.consumeDigits()
	}

	// Scientific notation: E or e, optionally followed by + or -.
	if l.pos < len(l.src) && (l.src[l.pos] == 'E' || l.src[l.pos] == 'e') {
		l.pos++
		if l.pos < len(l.src) && (l.src[l.pos] == '+' || l.src[l.pos] == '-') {
			l.pos++
		}
		l.consumeDigits()
	}

	return Token{Type: TokNumber, Value: string(l.src[start:l.pos]), Pos: start}, nil
}

// lexQuotedRef reads a quoted sheet reference: 'Sheet Name'!A1
// The token value is the full text including quotes and bang.
func (l *Lexer) lexQuotedRef() (Token, error) {
	start := l.pos
	l.pos++ // skip opening quote

	// Read until closing single quote. Doubled '' is an escape.
	for l.pos < len(l.src) {
		ch := l.src[l.pos]
		if ch == '\'' {
			l.pos++
			if l.pos < len(l.src) && l.src[l.pos] == '\'' {
				// Escaped quote, continue.
				l.pos++
				continue
			}
			// End of quoted name. Expect '!' next.
			if l.pos < len(l.src) && l.src[l.pos] == '!' {
				l.pos++ // skip !
				// Now read the cell/range reference part.
				refStart := l.pos
				l.consumeCellRefChars()
				if l.pos == refStart {
					return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
				}
				return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
			}
			return Token{}, fmt.Errorf("expected '!' after quoted sheet name at position %d", l.pos)
		}
		l.pos++
	}
	return Token{}, fmt.Errorf("unterminated quoted sheet name starting at position %d", start)
}

// lexIdentOrRef reads an identifier which could be:
//   - A cell reference: A1, $A$1
//   - A sheet-qualified reference: Sheet1!A1
//   - A function call: SUM(
//   - A boolean: TRUE, FALSE
func (l *Lexer) lexIdentOrRef() (Token, error) {
	start := l.pos

	// Handle $ prefix for absolute references.
	hasDollar := false
	if l.src[l.pos] == '$' {
		hasDollar = true
		l.pos++
	}

	// Read the alpha part.
	alphaStart := l.pos
	for l.pos < len(l.src) && isAlpha(l.src[l.pos]) {
		l.pos++
	}
	alpha := string(l.src[alphaStart:l.pos])

	// Check for pure-alpha sheet reference: Name!CellRef (e.g. Sheet!A1)
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == '!' {
		l.pos++ // skip !
		refStart := l.pos
		l.consumeCellRefChars()
		if l.pos == refStart {
			return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
		}
		return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
	}

	// Check for booleans (only if no $ prefix and pure alpha).
	if !hasDollar && len(alpha) > 0 {
		upper := strings.ToUpper(alpha)
		if upper == "TRUE" || upper == "FALSE" {
			// Make sure it's not followed by ( which would make it a function,
			// or alphanumeric which would make it part of a longer identifier.
			if l.pos >= len(l.src) || (!isAlphaNum(l.src[l.pos]) && l.src[l.pos] != '(' && l.src[l.pos] != '$') {
				return Token{Type: TokBool, Value: upper, Pos: start}, nil
			}
		}
	}

	// Skip optional $ before digits (for absolute row like A$1).
	if l.pos < len(l.src) && l.src[l.pos] == '$' {
		l.pos++
	}

	// Check for digits following the alpha part => cell reference or sheet name with digits.
	digitStart := l.pos
	for l.pos < len(l.src) && l.src[l.pos] >= '0' && l.src[l.pos] <= '9' {
		l.pos++
	}
	hasDigits := l.pos > digitStart

	// Check for sheet-qualified reference with digits in sheet name: Sheet1!A1
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == '!' {
		l.pos++ // skip !
		refStart := l.pos
		l.consumeCellRefChars()
		if l.pos == refStart {
			return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
		}
		return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
	}

	if hasDigits && len(alpha) > 0 {
		// Verify it's not followed by alphanumeric (which would mean it's an identifier, not a cell ref).
		// Also skip if followed by '(' — that makes it a function call (e.g. LOG10()).
		if l.pos >= len(l.src) || !isIdentContinue(l.src[l.pos]) {
			if l.pos >= len(l.src) || l.src[l.pos] != '(' {
				val := string(l.src[start:l.pos])
				if looksLikeCellRef(val) {
					return Token{Type: TokCellRef, Value: val, Pos: start}, nil
				}
			}
		}
	}

	// If we consumed a $ + digits but it's not a valid cell ref, backtrack.
	if hasDollar && len(alpha) == 0 {
		// Standalone $ is not valid.
		return Token{}, fmt.Errorf("unexpected '$' at position %d", start)
	}

	// Read remaining identifier chars (for function names or named ranges).
	for l.pos < len(l.src) && isIdentContinue(l.src[l.pos]) {
		l.pos++
	}

	// Check again for sheet-qualified ref after reading full identifier (e.g. Sheet_Name!A1).
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == '!' {
		l.pos++ // skip !
		refStart := l.pos
		l.consumeCellRefChars()
		if l.pos == refStart {
			return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
		}
		return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
	}

	word := string(l.src[start:l.pos])

	// Function call: identifier followed by '('.
	if l.pos < len(l.src) && l.src[l.pos] == '(' {
		l.pos++ // consume the '('
		return Token{Type: TokFunc, Value: word + "(", Pos: start}, nil
	}

	// Re-check booleans after full read (handles edge case of partial match above).
	upper := strings.ToUpper(word)
	if upper == "TRUE" || upper == "FALSE" {
		return Token{Type: TokBool, Value: upper, Pos: start}, nil
	}

	// If it looks like a cell ref, return as cell ref.
	if looksLikeCellRef(word) {
		return Token{Type: TokCellRef, Value: word, Pos: start}, nil
	}

	// Anything else is treated as a cell ref / named range.
	// The parser can disambiguate further.
	return Token{Type: TokCellRef, Value: word, Pos: start}, nil
}

// consumeCellRefChars advances past characters valid in a cell reference ($, letters, digits).
func (l *Lexer) consumeCellRefChars() {
	for l.pos < len(l.src) {
		ch := l.src[l.pos]
		if ch == '$' || isAlpha(ch) || (ch >= '0' && ch <= '9') {
			l.pos++
		} else {
			break
		}
	}
}

func (l *Lexer) consumeDigits() {
	for l.pos < len(l.src) && l.src[l.pos] >= '0' && l.src[l.pos] <= '9' {
		l.pos++
	}
}

func (l *Lexer) skipWhitespace() {
	for l.pos < len(l.src) && (l.src[l.pos] == ' ' || l.src[l.pos] == '\t' || l.src[l.pos] == '\n' || l.src[l.pos] == '\r') {
		l.pos++
	}
}

// looksLikeCellRef checks if a string looks like a valid Excel cell reference.
// Accepts: A1, $A1, A$1, $A$1, XFD1048576, etc.
func looksLikeCellRef(s string) bool {
	i := 0
	if i < len(s) && s[i] == '$' {
		i++
	}
	// Need at least one letter.
	alphaStart := i
	for i < len(s) && isAlpha(s[i]) {
		i++
	}
	if i == alphaStart {
		return false
	}
	// Column letters must be A-XFD (1-16384). Quick check: at most 3 letters.
	alphaLen := i - alphaStart
	if alphaLen > 3 {
		return false
	}
	if i < len(s) && s[i] == '$' {
		i++
	}
	// Need at least one digit.
	digitStart := i
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		i++
	}
	if i == digitStart {
		return false
	}
	// Must consume entire string.
	return i == len(s)
}

func isAlpha(ch byte) bool {
	return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')
}

func isAlphaNum(ch byte) bool {
	return isAlpha(ch) || (ch >= '0' && ch <= '9')
}

func isIdentStart(ch byte) bool {
	return isAlpha(ch) || ch == '_' || ch == '$' || ch == '\\'
}

func isIdentContinue(ch byte) bool {
	return isAlpha(ch) || (ch >= '0' && ch <= '9') || ch == '_' || ch == '.' || ch == '$'
}

func isExcelError(s string) bool {
	upper := strings.ToUpper(s)
	switch upper {
	case "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA":
		return true
	}
	return false
}
