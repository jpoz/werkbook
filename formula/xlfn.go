package formula

import "strings"

// xlfnPrefix maps function names to their required OOXML prefix.
// These are Excel functions added after the original OOXML specification
// that require a special prefix in the XML to be recognized by Excel.
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
	"ISOWEEKNUM":      "_xlfn.",
	"LAMBDA":          "_xlfn.",
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
