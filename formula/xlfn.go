package formula

import "strings"

// xlfnPrefix maps function names to their required OOXML prefix.
// These are Excel functions added after the original OOXML specification
// that require a special prefix in the XML to be recognized by Excel.
var xlfnPrefix = map[string]string{
	"ARABIC":      "_xlfn.",
	"BASE":        "_xlfn.",
	"COMBINA":     "_xlfn.",
	"CONCAT":      "_xlfn.",
	"COT":         "_xlfn.",
	"COTH":        "_xlfn.",
	"CSC":         "_xlfn.",
	"CSCH":        "_xlfn.",
	"DAYS":        "_xlfn.",
	"DECIMAL":     "_xlfn.",
	"IFNA":        "_xlfn.",
	"IFS":         "_xlfn.",
	"ISOWEEKNUM":  "_xlfn.",
	"MAXIFS":      "_xlfn.",
	"MINIFS":      "_xlfn.",
	"NUMBERVALUE": "_xlfn.",
	"SEC":         "_xlfn.",
	"SECH":        "_xlfn.",
	"SORT":        "_xlfn._xlws.",
	"SWITCH":      "_xlfn.",
	"TEXTJOIN":    "_xlfn.",
	"XLOOKUP":     "_xlfn.",
	"XOR":         "_xlfn.",
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
