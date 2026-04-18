package formula

import (
	"fmt"
	"strings"
)

// Lexer tokenizes a formula string.
type Lexer struct {
	src     []byte
	pos     int       // current byte position
	prev    TokenType // last emitted token type (TokEOF when none)
	prevVal string    // last emitted token value (used for disambiguation)
	pendInt bool      // emit a synthetic TokIntersect before the next token
	pendPos int       // position of the whitespace that introduced pendInt
	buf     Token     // one-token lookahead buffer used when injecting TokIntersect
	hasBuf  bool
}

// NewLexer creates a lexer for the given formula string.
// The formula should NOT include the leading '=';
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
	tok, err := l.nextRaw()
	if err != nil {
		return tok, err
	}
	// Apply pending intersection detection: a whitespace gap between the
	// previously emitted token and this one becomes a TokIntersect only when
	// both sides are reference-like.
	if l.pendInt && canPrecedeIntersect(l.prev, l.prevVal) && canFollowIntersect(tok.Type, tok.Value) {
		intersectTok := Token{Type: TokIntersect, Value: " ", Pos: l.pendPos}
		l.pendInt = false
		l.pendPos = 0
		l.buf = tok
		l.hasBuf = true
		l.prev = intersectTok.Type
		l.prevVal = intersectTok.Value
		return intersectTok, nil
	}
	l.pendInt = false
	l.pendPos = 0
	l.prev = tok.Type
	l.prevVal = tok.Value
	return tok, nil
}

// nextRaw reads the next token from the input stream, detecting leading
// whitespace and recording it as a potential intersection hint via pendInt.
func (l *Lexer) nextRaw() (Token, error) {
	if l.hasBuf {
		t := l.buf
		l.hasBuf = false
		l.buf = Token{}
		return t, nil
	}
	l.skipWhitespaceTracking()

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
			if isFormulaError(val) {
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
	if isFormulaError(val) {
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
	// Also check for 3D reference: Name:Name!CellRef (e.g. Sheet:Sheet5!A1)
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == '!' {
		l.pos++ // skip !
		refStart := l.pos
		l.consumeCellRefChars()
		if l.pos == refStart {
			return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
		}
		return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
	}
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == ':' {
		if bangPos := l.find3DSheetBang(); bangPos > 0 {
			l.pos = bangPos + 1 // skip past !
			refStart := l.pos
			l.consumeCellRefChars()
			if l.pos == refStart {
				return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
			}
			return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
		}
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
	// Also check for 3D reference: Sheet1:Sheet5!A1
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == '!' {
		l.pos++ // skip !
		refStart := l.pos
		l.consumeCellRefChars()
		if l.pos == refStart {
			return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
		}
		return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
	}
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == ':' {
		// Only try 3D sheet reference if the word so far does NOT look like
		// a cell reference. E.g. "S1:S3!A1" — S1 is a valid cell ref, so
		// the colon should be a range operator, not a 3D sheet separator.
		// Sheet names that look like cell refs
		// must be quoted to form a valid 3D reference.
		word := string(l.src[start:l.pos])
		if !looksLikeCellRef(word) {
			if bangPos := l.find3DSheetBang(); bangPos > 0 {
				l.pos = bangPos + 1 // skip past !
				refStart := l.pos
				l.consumeCellRefChars()
				if l.pos == refStart {
					return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
				}
				return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
			}
		}
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
	if !hasDollar && l.pos < len(l.src) && l.src[l.pos] == ':' {
		word := string(l.src[start:l.pos])
		if !looksLikeCellRef(word) {
			if bangPos := l.find3DSheetBang(); bangPos > 0 {
				l.pos = bangPos + 1 // skip past !
				refStart := l.pos
				l.consumeCellRefChars()
				if l.pos == refStart {
					return Token{}, fmt.Errorf("expected cell reference after '!' at position %d", l.pos)
				}
				return Token{Type: TokCellRef, Value: string(l.src[start:l.pos]), Pos: start}, nil
			}
		}
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

// find3DSheetBang checks whether the lexer is sitting on a ':' that is part of
// a 3D sheet reference (e.g. Sheet2:Sheet5!A1). If so, it returns the position
// of the '!' character. If not, it returns -1 and does not advance the lexer.
func (l *Lexer) find3DSheetBang() int {
	if l.pos >= len(l.src) || l.src[l.pos] != ':' {
		return -1
	}
	// Scan past ':' and look for an identifier followed by '!'.
	p := l.pos + 1
	// The second sheet name can contain letters, digits, underscores, dots.
	nameStart := p
	for p < len(l.src) && isIdentContinue(l.src[p]) {
		p++
	}
	if p == nameStart {
		return -1
	}
	if p < len(l.src) && l.src[p] == '!' {
		return p
	}
	return -1
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

// skipWhitespaceTracking skips whitespace like skipWhitespace but records a
// pending intersect hint if any whitespace was consumed.
func (l *Lexer) skipWhitespaceTracking() {
	start := l.pos
	for l.pos < len(l.src) && (l.src[l.pos] == ' ' || l.src[l.pos] == '\t' || l.src[l.pos] == '\n' || l.src[l.pos] == '\r') {
		l.pos++
	}
	if l.pos > start {
		l.pendInt = true
		l.pendPos = start
	}
}

// canPrecedeIntersect reports whether a token of the given type/value can act
// as the left operand of an intersection (i.e. produces a reference).
func canPrecedeIntersect(t TokenType, _ string) bool {
	switch t {
	case TokCellRef, TokRParen:
		return true
	}
	return false
}

// canFollowIntersect reports whether a token of the given type/value can act
// as the right operand of an intersection (i.e. produces a reference).
func canFollowIntersect(t TokenType, val string) bool {
	switch t {
	case TokCellRef, TokLParen:
		return true
	case TokFunc:
		// Only reference-returning functions can participate in an
		// intersection. We accept a small curated set; anything else
		// causes the space to be ignored (Excel behavior).
		switch strings.ToUpper(strings.TrimSuffix(val, "(")) {
		case "OFFSET", "INDIRECT", "INDEX", "CHOOSE", "IF", "IFS", "SWITCH":
			return true
		}
	}
	return false
}

// looksLikeCellRef checks if a string looks like a valid cell reference.
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

func isFormulaError(s string) bool {
	upper := strings.ToUpper(s)
	switch upper {
	case "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA":
		return true
	}
	return false
}
