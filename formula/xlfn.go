package formula

import "strings"

const xlpmPrefix = "_xlpm."

type formulaInsertion struct {
	pos    int
	prefix string
}

// xlfnPrefix maps function names to their required OOXML prefix.
// These are Excel functions added after the original OOXML specification
// that require a special prefix in the XML to be recognized by Excel.
var xlfnPrefix = map[string]string{
	"ANCHORARRAY":     "_xlfn.",
	"ARABIC":          "_xlfn.",
	"BASE":            "_xlfn.",
	"BYCOL":           "_xlfn.",
	"BYROW":           "_xlfn.",
	"CEILING.MATH":    "_xlfn.",
	"CEILING.PRECISE": "_xlfn.",
	"CHOOSECOLS":      "_xlfn.",
	"CHOOSEROWS":      "_xlfn.",
	"COMBINA":         "_xlfn.",
	"CONCAT":          "_xlfn.",
	"COT":             "_xlfn.",
	"COTH":            "_xlfn.",
	"COVARIANCE.P":    "_xlfn.",
	"COVARIANCE.S":    "_xlfn.",
	"CSC":             "_xlfn.",
	"CSCH":            "_xlfn.",
	"DAYS":            "_xlfn.",
	"DECIMAL":         "_xlfn.",
	"DROP":            "_xlfn.",
	"EXPAND":          "_xlfn.",
	"FILTER":          "_xlfn._xlws.",
	"FLOOR.MATH":      "_xlfn.",
	"FLOOR.PRECISE":   "_xlfn.",
	"FORECAST.LINEAR": "_xlfn.",
	"GAMMALN.PRECISE": "_xlfn.",
	"HSTACK":          "_xlfn.",
	"IFNA":            "_xlfn.",
	"IFS":             "_xlfn.",
	"ISOWEEKNUM":      "_xlfn.",
	"LET":             "_xlfn.",
	"LAMBDA":          "_xlfn.",
	"MAKEARRAY":       "_xlfn.",
	"MAP":             "_xlfn.",
	"MAXIFS":          "_xlfn.",
	"MINIFS":          "_xlfn.",
	"MODE.SNGL":       "_xlfn.",
	"NORM.DIST":       "_xlfn.",
	"NORM.INV":        "_xlfn.",
	"NORM.S.DIST":     "_xlfn.",
	"NORM.S.INV":      "_xlfn.",
	"NUMBERVALUE":     "_xlfn.",
	"PERCENTILE.EXC":  "_xlfn.",
	"PERCENTRANK.EXC": "_xlfn.",
	"PERCENTRANK.INC": "_xlfn.",
	"RANDARRAY":       "_xlfn.",
	"RANK.AVG":        "_xlfn.",
	"RANK.EQ":         "_xlfn.",
	"QUARTILE.EXC":    "_xlfn.",
	"REDUCE":          "_xlfn.",
	"SCAN":            "_xlfn.",
	"SEC":             "_xlfn.",
	"SECH":            "_xlfn.",
	"SEQUENCE":        "_xlfn.",
	"SINGLE":          "_xlfn.",
	"SORT":            "_xlfn._xlws.",
	"SORTBY":          "_xlfn.",
	"STDEV.P":         "_xlfn.",
	"STDEV.S":         "_xlfn.",
	"SWITCH":          "_xlfn.",
	"TAKE":            "_xlfn.",
	"TEXTJOIN":        "_xlfn.",
	"TEXTSPLIT":       "_xlfn.",
	"TOCOL":           "_xlfn.",
	"TOROW":           "_xlfn.",
	"UNIQUE":          "_xlfn.",
	"VAR.P":           "_xlfn.",
	"VAR.S":           "_xlfn.",
	"VSTACK":          "_xlfn.",
	"WRAPCOLS":        "_xlfn.",
	"WRAPROWS":        "_xlfn.",
	"XLOOKUP":         "_xlfn.",
	"XOR":             "_xlfn.",
}

var dynamicArrayFunctions = map[string]struct{}{
	"ANCHORARRAY": {},
	"BYCOL":       {},
	"BYROW":       {},
	"CHOOSECOLS":  {},
	"CHOOSEROWS":  {},
	"DROP":        {},
	"EXPAND":      {},
	"FILTER":      {},
	"HSTACK":      {},
	"LAMBDA":      {},
	"MAKEARRAY":   {},
	"MAP":         {},
	"RANDARRAY":   {},
	"REDUCE":      {},
	"SCAN":        {},
	"SEQUENCE":    {},
	"SINGLE":      {},
	"SORT":        {},
	"SORTBY":      {},
	"SWITCH":      {},
	"TAKE":        {},
	"TEXTSPLIT":   {},
	"TOCOL":       {},
	"TOROW":       {},
	"UNIQUE":      {},
	"VSTACK":      {},
	"WRAPCOLS":    {},
	"WRAPROWS":    {},
	"XLOOKUP":     {},
}

// AddXlfnPrefixes tokenizes the formula and inserts the required OOXML
// prefixes (e.g. _xlfn.) before function names that need them. It only
// modifies TokFunc tokens, so strings and cell references are never touched.
// Returns the original string unchanged if tokenization fails.
func AddXlfnPrefixes(f string) string {
	if f == "" {
		return f
	}
	f = addLETParamPrefixes(f)
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	// Collect insertions (position + prefix string) from right to left
	// so earlier positions are unaffected by later splicing.
	var inserts []formulaInsertion

	for _, tok := range tokens {
		if tok.Type != TokFunc {
			continue
		}
		// TokFunc value includes the trailing '(' — strip it for lookup.
		name := strings.ToUpper(strings.TrimSuffix(tok.Value, "("))
		// Already prefixed (e.g. round-trip) — skip.
		if strings.HasPrefix(name, "_XLFN.") {
			continue
		}
		prefix, ok := xlfnPrefix[name]
		if !ok {
			continue
		}
		inserts = append(inserts, formulaInsertion{pos: tok.Pos, prefix: prefix})
	}

	if len(inserts) == 0 {
		return f
	}

	// Apply from right to left so byte offsets stay valid.
	buf := []byte(f)
	for i := len(inserts) - 1; i >= 0; i-- {
		ins := inserts[i]
		p := []byte(ins.prefix)
		buf = append(buf[:ins.pos], append(p, buf[ins.pos:]...)...)
	}
	return string(buf)
}

// IsDynamicArrayFormula reports whether the formula uses Excel dynamic-array
// semantics and must be serialized with dynamic-array OOXML metadata.
func IsDynamicArrayFormula(f string) bool {
	if f == "" {
		return false
	}
	tokens, err := Tokenize(f)
	if err == nil {
		for _, tok := range tokens {
			if tok.Type != TokFunc {
				continue
			}
			name := normalizeFuncName(strings.TrimSuffix(tok.Value, "("))
			if _, ok := dynamicArrayFunctions[name]; ok {
				return true
			}
		}
		return false
	}

	upper := strings.ToUpper(f)
	for name := range dynamicArrayFunctions {
		if strings.Contains(upper, name+"(") ||
			strings.Contains(upper, "_XLFN."+name+"(") ||
			strings.Contains(upper, "_XLFN._XLWS."+name+"(") {
			return true
		}
	}
	return false
}

// StripXlfnPrefixes tokenizes the formula and removes OOXML prefixes
// (_xlfn. and _xlfn._xlws.) from function names. Only TokFunc tokens
// are modified, so strings and cell references are never touched.
// Returns the original string unchanged if tokenization fails.
func StripXlfnPrefixes(f string) string {
	if f == "" {
		return f
	}
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	type removal struct {
		pos int
		len int
	}
	var removals []removal

	for _, tok := range tokens {
		if tok.Type != TokFunc {
			continue
		}
		upper := strings.ToUpper(tok.Value)
		switch {
		case strings.HasPrefix(upper, "_XLFN._XLWS."):
			removals = append(removals, removal{pos: tok.Pos, len: len("_xlfn._xlws.")})
		case strings.HasPrefix(upper, "_XLFN."):
			removals = append(removals, removal{pos: tok.Pos, len: len("_xlfn.")})
		}
	}

	if len(removals) == 0 {
		return stripXLPMPrefixes(f)
	}

	// Apply from right to left so byte offsets stay valid.
	buf := []byte(f)
	for i := len(removals) - 1; i >= 0; i-- {
		r := removals[i]
		buf = append(buf[:r.pos], buf[r.pos+r.len:]...)
	}
	return stripXLPMPrefixes(string(buf))
}

func normalizeFuncName(name string) string {
	upper := strings.ToUpper(name)
	upper = strings.TrimPrefix(upper, "_XLFN._XLWS.")
	upper = strings.TrimPrefix(upper, "_XLFN.")
	return upper
}

func addLETParamPrefixes(f string) string {
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	var inserts []formulaInsertion

	walkLETTokenRange(tokens, 0, len(tokens)-1, map[string]int{}, &inserts)
	if len(inserts) == 0 {
		return f
	}

	buf := []byte(f)
	for i := len(inserts) - 1; i >= 0; i-- {
		ins := inserts[i]
		buf = append(buf[:ins.pos], append([]byte(ins.prefix), buf[ins.pos:]...)...)
	}
	return string(buf)
}

func walkLETTokenRange(tokens []Token, start, end int, scope map[string]int, inserts *[]formulaInsertion) {
	for i := start; i < end; i++ {
		tok := tokens[i]
		switch tok.Type {
		case TokFunc:
			close := findFuncClose(tokens, i)
			if close < 0 || close > end {
				return
			}
			name := normalizeFuncName(strings.TrimSuffix(tok.Value, "("))
			if name == "LET" {
				processLETTokens(tokens, i+1, close, scope, inserts)
			} else {
				walkLETTokenRange(tokens, i+1, close, scope, inserts)
			}
			i = close
		case TokCellRef:
			if shouldPrefixLETName(tokens, i, scope) {
				*inserts = append(*inserts, formulaInsertion{pos: tok.Pos, prefix: xlpmPrefix})
			}
		}
	}
}

func processLETTokens(tokens []Token, start, end int, parent map[string]int, inserts *[]formulaInsertion) {
	local := cloneLETNameScope(parent)
	args := splitFuncArgRanges(tokens, start, end)
	last := len(args) - 1

	for i, arg := range args {
		isLast := i == last
		if !isLast && i%2 == 0 {
			tok, ok := singleLETNameToken(tokens[arg.start:arg.end])
			if !ok {
				continue
			}
			if !hasXLPMPrefix(tok.Value) {
				*inserts = append(*inserts, formulaInsertion{pos: tok.Pos, prefix: xlpmPrefix})
			}
			local[strings.ToUpper(stripXLPMNamePrefix(tok.Value))]++
			continue
		}
		walkLETTokenRange(tokens, arg.start, arg.end, local, inserts)
	}
}

func cloneLETNameScope(scope map[string]int) map[string]int {
	out := make(map[string]int, len(scope))
	for k, v := range scope {
		out[k] = v
	}
	return out
}

type tokenRange struct {
	start int
	end   int
}

func splitFuncArgRanges(tokens []Token, start, end int) []tokenRange {
	if start >= end {
		return nil
	}
	var out []tokenRange
	argStart := start
	depth := 0
	arrayDepth := 0

	for i := start; i < end; i++ {
		switch tokens[i].Type {
		case TokFunc, TokLParen:
			depth++
		case TokRParen:
			if depth > 0 {
				depth--
			}
		case TokArrayOpen:
			arrayDepth++
		case TokArrayClose:
			if arrayDepth > 0 {
				arrayDepth--
			}
		case TokComma:
			if depth == 0 && arrayDepth == 0 {
				out = append(out, tokenRange{start: argStart, end: i})
				argStart = i + 1
			}
		}
	}

	out = append(out, tokenRange{start: argStart, end: end})
	return out
}

func findFuncClose(tokens []Token, funcIdx int) int {
	depth := 0
	for i := funcIdx + 1; i < len(tokens); i++ {
		switch tokens[i].Type {
		case TokFunc, TokLParen:
			depth++
		case TokRParen:
			if depth == 0 {
				return i
			}
			depth--
		}
	}
	return -1
}

func singleLETNameToken(tokens []Token) (Token, bool) {
	if len(tokens) != 1 || tokens[0].Type != TokCellRef {
		return Token{}, false
	}
	name := stripXLPMNamePrefix(tokens[0].Value)
	if !isValidLETName(name) {
		return Token{}, false
	}
	return tokens[0], true
}

func shouldPrefixLETName(tokens []Token, idx int, scope map[string]int) bool {
	tok := tokens[idx]
	name := stripXLPMNamePrefix(tok.Value)
	if !isBareNameToken(name) || hasXLPMPrefix(tok.Value) {
		return false
	}
	if idx+1 < len(tokens) && tokens[idx+1].Type == TokColon {
		return false
	}
	return scope[strings.ToUpper(name)] > 0
}

func stripXLPMPrefixes(f string) string {
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	type removal struct {
		pos int
		len int
	}
	var removals []removal
	for _, tok := range tokens {
		if tok.Type != TokCellRef || !hasXLPMPrefix(tok.Value) {
			continue
		}
		removals = append(removals, removal{pos: tok.Pos, len: len(xlpmPrefix)})
	}
	if len(removals) == 0 {
		return f
	}

	buf := []byte(f)
	for i := len(removals) - 1; i >= 0; i-- {
		r := removals[i]
		buf = append(buf[:r.pos], buf[r.pos+r.len:]...)
	}
	return string(buf)
}

func hasXLPMPrefix(name string) bool {
	return strings.HasPrefix(strings.ToUpper(name), strings.ToUpper(xlpmPrefix))
}

func stripXLPMNamePrefix(name string) string {
	if hasXLPMPrefix(name) {
		return name[len(xlpmPrefix):]
	}
	return name
}
