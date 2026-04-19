package formula

import "strings"

// xlfnPrefix maps function names to their required OOXML prefix.
// These are functions added after the original OOXML specification
// that require a special prefix in the XML to be recognized.
var xlfnPrefix = map[string]string{
	"ACOT":            "_xlfn.",
	"ACOTH":           "_xlfn.",
	"ANCHORARRAY":     "_xlfn.",
	"ARABIC":          "_xlfn.",
	"BASE":            "_xlfn.",
	"BETA.DIST":       "_xlfn.",
	"BETA.INV":        "_xlfn.",
	"BITAND":          "_xlfn.",
	"BITLSHIFT":       "_xlfn.",
	"BITOR":           "_xlfn.",
	"BITRSHIFT":       "_xlfn.",
	"BITXOR":          "_xlfn.",
	"BINOM.DIST":      "_xlfn.",
	"BINOM.INV":       "_xlfn.",
	"BYCOL":           "_xlfn.",
	"BYROW":           "_xlfn.",
	"CHISQ.DIST":      "_xlfn.",
	"CHISQ.DIST.RT":   "_xlfn.",
	"CHISQ.INV":       "_xlfn.",
	"CHISQ.INV.RT":    "_xlfn.",
	"CHISQ.TEST":      "_xlfn.",
	"CEILING.MATH":    "_xlfn.",
	"CEILING.PRECISE": "_xlfn.",
	"CONFIDENCE.NORM": "_xlfn.",
	"CONFIDENCE.T":    "_xlfn.",
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
	"ERF.PRECISE":     "_xlfn.",
	"ERFC.PRECISE":    "_xlfn.",
	"EXPON.DIST":      "_xlfn.",
	"EXPAND":          "_xlfn.",
	"F.DIST":          "_xlfn.",
	"F.DIST.RT":       "_xlfn.",
	"F.INV":           "_xlfn.",
	"F.INV.RT":        "_xlfn.",
	"FILTER":          "_xlfn._xlws.",
	"FLOOR.MATH":      "_xlfn.",
	"FLOOR.PRECISE":   "_xlfn.",
	"FORECAST.LINEAR": "_xlfn.",
	"GAUSS":           "_xlfn.",
	"GAMMA":           "_xlfn.",
	"GAMMA.DIST":      "_xlfn.",
	"GAMMA.INV":       "_xlfn.",
	"GAMMALN.PRECISE": "_xlfn.",
	"HSTACK":          "_xlfn.",
	"HYPGEOM.DIST":    "_xlfn.",
	"IFNA":            "_xlfn.",
	"IFS":             "_xlfn.",
	"IMCOSH":          "_xlfn.",
	"IMCOT":           "_xlfn.",
	"IMCSC":           "_xlfn.",
	"IMCSCH":          "_xlfn.",
	"IMSECH":          "_xlfn.",
	"IMSINH":          "_xlfn.",
	"ISOMITTED":       "_xlfn.",
	"ISOWEEKNUM":      "_xlfn.",
	"LAMBDA":          "_xlfn.",
	"LET":             "_xlfn.",
	"MAKEARRAY":       "_xlfn.",
	"MAP":             "_xlfn.",
	"MAXIFS":          "_xlfn.",
	"MINIFS":          "_xlfn.",
	"MODE.SNGL":       "_xlfn.",
	"NEGBINOM.DIST":   "_xlfn.",
	"LOGNORM.DIST":    "_xlfn.",
	"LOGNORM.INV":     "_xlfn.",
	"NORM.DIST":       "_xlfn.",
	"NORM.INV":        "_xlfn.",
	"NORM.S.DIST":     "_xlfn.",
	"NORM.S.INV":      "_xlfn.",
	"NUMBERVALUE":     "_xlfn.",
	"PERCENTILE.EXC":  "_xlfn.",
	"PERCENTILE.INC":  "_xlfn.",
	"PHI":             "_xlfn.",
	"PDURATION":       "_xlfn.",
	"POISSON.DIST":    "_xlfn.",
	"PERCENTRANK.EXC": "_xlfn.",
	"PERCENTRANK.INC": "_xlfn.",
	"RANDARRAY":       "_xlfn.",
	"RANK.AVG":        "_xlfn.",
	"RANK.EQ":         "_xlfn.",
	"QUARTILE.EXC":    "_xlfn.",
	"QUARTILE.INC":    "_xlfn.",
	"REDUCE":          "_xlfn.",
	"RRI":             "_xlfn.",
	"SCAN":            "_xlfn.",
	"SEC":             "_xlfn.",
	"SECH":            "_xlfn.",
	"SEQUENCE":        "_xlfn.",
	"SINGLE":          "_xlfn.",
	"SORT":            "_xlfn._xlws.",
	"SORTBY":          "_xlfn.",
	"SKEW.P":          "_xlfn.",
	"STDEV.P":         "_xlfn.",
	"STDEV.S":         "_xlfn.",
	"T.DIST":          "_xlfn.",
	"T.DIST.2T":       "_xlfn.",
	"T.DIST.RT":       "_xlfn.",
	"T.INV":           "_xlfn.",
	"T.INV.2T":        "_xlfn.",
	"T.TEST":          "_xlfn.",
	"SWITCH":          "_xlfn.",
	"TAKE":            "_xlfn.",
	"TEXTJOIN":        "_xlfn.",
	"TEXTSPLIT":       "_xlfn.",
	"TOCOL":           "_xlfn.",
	"TOROW":           "_xlfn.",
	"UNICHAR":         "_xlfn.",
	"UNICODE":         "_xlfn.",
	"UNIQUE":          "_xlfn.",
	"VAR.P":           "_xlfn.",
	"VAR.S":           "_xlfn.",
	"VSTACK":          "_xlfn.",
	"WEIBULL.DIST":    "_xlfn.",
	"WRAPCOLS":        "_xlfn.",
	"WRAPROWS":        "_xlfn.",
	"XLOOKUP":         "_xlfn.",
	"XMATCH":          "_xlfn.",
	"XOR":             "_xlfn.",
	"Z.TEST":          "_xlfn.",
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
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	// Collect insertions (position + prefix string) from right to left
	// so earlier positions are unaffected by later splicing.
	type insertion struct {
		pos    int
		prefix string
	}
	var inserts []insertion

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
		inserts = append(inserts, insertion{pos: tok.Pos, prefix: prefix})
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

// AddXlpmPrefixes tokenizes the formula and inserts the required OOXML
// _xlpm. prefix on LET/LAMBDA parameter names and references. This matches
// how Excel serializes future formulas in workbook XML, for example:
//
//	LET(x,5,x+1) -> _xlfn.LET(_xlpm.x,5,_xlpm.x+1)
//	MAP(A1:A3,LAMBDA(x,x+1)) -> _xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1))
//
// Returns the original string unchanged if tokenization fails.
func AddXlpmPrefixes(f string) string {
	if f == "" {
		return f
	}
	tokens, err := Tokenize(f)
	if err != nil {
		return f
	}

	var inserts []int
	collectXlpmInsertions(tokens, 0, len(tokens)-1, nil, &inserts)
	if len(inserts) == 0 {
		return f
	}

	buf := []byte(f)
	for i := len(inserts) - 1; i >= 0; i-- {
		pos := inserts[i]
		buf = append(buf[:pos], append([]byte("_xlpm."), buf[pos:]...)...)
	}
	return string(buf)
}

// IsDynamicArrayFormula reports whether the formula uses dynamic-array
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

// StripXlfnPrefixes tokenizes the formula and removes OOXML prefixes used in
// workbook XML (_xlfn., _xlfn._xlws., and _xlpm.). Only formula/function and
// lambda-parameter identifier tokens are modified, so strings are never touched.
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
		switch tok.Type {
		case TokFunc:
			upper := strings.ToUpper(tok.Value)
			switch {
			case strings.HasPrefix(upper, "_XLFN._XLWS."):
				removals = append(removals, removal{pos: tok.Pos, len: len("_xlfn._xlws.")})
			case strings.HasPrefix(upper, "_XLFN."):
				removals = append(removals, removal{pos: tok.Pos, len: len("_xlfn.")})
			case strings.HasPrefix(upper, "_XLPM."):
				// Workbook-level named LAMBDAs and LET-bound LAMBDAs are
				// serialized as `_xlpm.name(args)` call-sites by Excel.
				removals = append(removals, removal{pos: tok.Pos, len: len("_xlpm.")})
			}
		case TokCellRef:
			upper := strings.ToUpper(tok.Value)
			if strings.HasPrefix(upper, "_XLPM.") {
				removals = append(removals, removal{pos: tok.Pos, len: len("_xlpm.")})
			}
		}
	}

	if len(removals) == 0 {
		return f
	}

	// Apply from right to left so byte offsets stay valid.
	buf := []byte(f)
	for i := len(removals) - 1; i >= 0; i-- {
		r := removals[i]
		buf = append(buf[:r.pos], buf[r.pos+r.len:]...)
	}
	return string(buf)
}

func normalizeFuncName(name string) string {
	upper := strings.ToUpper(name)
	upper = strings.TrimPrefix(upper, "_XLFN._XLWS.")
	upper = strings.TrimPrefix(upper, "_XLFN.")
	return upper
}

type tokenSpan struct {
	start int
	end   int
}

// Lambda and LET parameter names are tokenized as TokCellRef because bare
// identifiers share the same lexical space as column-only references. This
// walk reinterprets the tokens based on function argument position and inserts
// _xlpm. only when the identifier is in scope as a lambda parameter.
func collectXlpmInsertions(tokens []Token, start, end int, scope map[string]struct{}, inserts *[]int) {
	for i := start; i < end; {
		tok := tokens[i]
		if tok.Type == TokFunc {
			next := processXlpmFunction(tokens, i, scope, inserts)
			if next > i {
				i = next
				continue
			}
		}
		if tok.Type == TokCellRef {
			if name, ok := xlpmIdentifierName(tok.Value); ok && scopeContains(scope, name) && !hasXlpmPrefix(tok.Value) {
				*inserts = append(*inserts, tok.Pos)
			}
		}
		i++
	}
}

func processXlpmFunction(tokens []Token, funcIdx int, scope map[string]struct{}, inserts *[]int) int {
	args, end, ok := splitFunctionArgs(tokens, funcIdx)
	if !ok {
		return funcIdx + 1
	}

	name := normalizeFuncName(strings.TrimSuffix(tokens[funcIdx].Value, "("))
	switch name {
	case "LET":
		local := cloneNameSet(scope)
		last := len(args) - 1
		if last < 0 {
			return end
		}
		for i := 0; i+1 < len(args); i += 2 {
			paramName, ok := markXlpmParam(tokens, args[i], inserts)
			if !ok {
				collectXlpmInsertions(tokens, args[i].start, args[i].end, local, inserts)
			}
			collectXlpmInsertions(tokens, args[i+1].start, args[i+1].end, local, inserts)
			if ok {
				local[paramName] = struct{}{}
			}
		}
		// LET's final argument is the calculation body when the arg count is odd.
		if len(args)%2 == 1 {
			collectXlpmInsertions(tokens, args[last].start, args[last].end, local, inserts)
		}
	case "LAMBDA":
		local := cloneNameSet(scope)
		last := len(args) - 1
		if last < 0 {
			return end
		}
		for i := 0; i < last; i++ {
			paramName, ok := markXlpmParam(tokens, args[i], inserts)
			if !ok {
				collectXlpmInsertions(tokens, args[i].start, args[i].end, scope, inserts)
				continue
			}
			local[paramName] = struct{}{}
		}
		collectXlpmInsertions(tokens, args[last].start, args[last].end, local, inserts)
	default:
		for _, arg := range args {
			collectXlpmInsertions(tokens, arg.start, arg.end, scope, inserts)
		}
	}

	return end
}

func splitFunctionArgs(tokens []Token, funcIdx int) ([]tokenSpan, int, bool) {
	var (
		args       []tokenSpan
		argStart   = funcIdx + 1
		parenDepth = 1
		arrayDepth int
	)
	for i := funcIdx + 1; i < len(tokens); i++ {
		switch tokens[i].Type {
		case TokFunc, TokLParen:
			parenDepth++
		case TokRParen:
			parenDepth--
			if parenDepth == 0 {
				args = append(args, tokenSpan{start: argStart, end: i})
				return args, i + 1, true
			}
		case TokArrayOpen:
			arrayDepth++
		case TokArrayClose:
			if arrayDepth > 0 {
				arrayDepth--
			}
		case TokComma:
			if parenDepth == 1 && arrayDepth == 0 {
				args = append(args, tokenSpan{start: argStart, end: i})
				argStart = i + 1
			}
		}
	}
	return nil, funcIdx + 1, false
}

func markXlpmParam(tokens []Token, span tokenSpan, inserts *[]int) (string, bool) {
	if span.end-span.start != 1 {
		return "", false
	}
	tok := tokens[span.start]
	if tok.Type != TokCellRef {
		return "", false
	}
	name, ok := xlpmIdentifierName(tok.Value)
	if !ok {
		return "", false
	}
	if !hasXlpmPrefix(tok.Value) {
		*inserts = append(*inserts, tok.Pos)
	}
	return name, true
}

func xlpmIdentifierName(raw string) (string, bool) {
	raw = stripXlpmPrefix(raw)
	ref, err := parseCellRefToken(raw)
	if err != nil || ref.Row != 0 || ref.Sheet != "" || ref.SheetEnd != "" || ref.AbsCol || ref.AbsRow || ref.DotNotation {
		return "", false
	}
	return strings.ToUpper(ColNumberToLetters(ref.Col)), true
}

func stripXlpmPrefix(raw string) string {
	if len(raw) >= len("_xlpm.") && strings.EqualFold(raw[:len("_xlpm.")], "_xlpm.") {
		return raw[len("_xlpm."):]
	}
	return raw
}

func hasXlpmPrefix(raw string) bool {
	return len(raw) >= len("_xlpm.") && strings.EqualFold(raw[:len("_xlpm.")], "_xlpm.")
}

func cloneNameSet(src map[string]struct{}) map[string]struct{} {
	if len(src) == 0 {
		return map[string]struct{}{}
	}
	dst := make(map[string]struct{}, len(src))
	for name := range src {
		dst[name] = struct{}{}
	}
	return dst
}

func scopeContains(scope map[string]struct{}, name string) bool {
	if len(scope) == 0 {
		return false
	}
	_, ok := scope[name]
	return ok
}
