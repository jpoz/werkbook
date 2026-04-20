package formula

import (
	"regexp"
	"strings"
	"sync"
	"unicode/utf8"
)

func init() {
	Register("REGEXTEST", NoCtx(fnREGEXTEST))
	Register("REGEXEXTRACT", NoCtx(fnREGEXEXTRACT))
	Register("REGEXREPLACE", NoCtx(fnREGEXREPLACE))
}

// regexCacheEntry holds either a compiled pattern or the compile error as a
// formula Value. Both are cached so invalid patterns don't retry compilation.
type regexCacheEntry struct {
	re  *regexp.Regexp
	err *Value
}

var regexCache sync.Map // map[string]regexCacheEntry, key = "0:pat" or "1:pat"

// compileRegex compiles pattern (optionally case-insensitive) and caches the
// result. Returns the regexp, or a formula error Value when the pattern is
// invalid (and the regexp is nil).
func compileRegex(pattern string, caseInsensitive bool) (*regexp.Regexp, *Value) {
	key := "0:" + pattern
	if caseInsensitive {
		key = "1:" + pattern
	}
	if v, ok := regexCache.Load(key); ok {
		e := v.(regexCacheEntry)
		return e.re, e.err
	}
	src := pattern
	if caseInsensitive {
		src = "(?i)" + pattern
	}
	re, err := regexp.Compile(src)
	if err != nil {
		v := ErrorVal(ErrValVALUE)
		entry := regexCacheEntry{err: &v}
		regexCache.Store(key, entry)
		return nil, &v
	}
	entry := regexCacheEntry{re: re}
	regexCache.Store(key, entry)
	return re, nil
}

// coerceRegexCaseFlag reads an Excel-style case flag: 0 = case-sensitive
// (default), non-zero = case-insensitive. Empty values treat as default.
func coerceRegexCaseFlag(v Value) (bool, *Value) {
	if v.Type == ValueEmpty {
		return false, nil
	}
	n, e := CoerceNum(v)
	if e != nil {
		return false, e
	}
	return n != 0, nil
}

func fnREGEXTEST(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	if args[1].Type == ValueError {
		return args[1], nil
	}
	caseInsensitive := false
	if len(args) == 3 {
		if args[2].Type == ValueError {
			return args[2], nil
		}
		flag, e := coerceRegexCaseFlag(args[2])
		if e != nil {
			return *e, nil
		}
		caseInsensitive = flag
	}
	re, errVal := compileRegex(ValueToString(args[1]), caseInsensitive)
	if errVal != nil {
		return *errVal, nil
	}
	return BoolVal(re.MatchString(ValueToString(args[0]))), nil
}

func fnREGEXEXTRACT(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[1].Type == ValueError {
		return args[1], nil
	}
	returnMode := 0
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		if args[2].Type == ValueError {
			return args[2], nil
		}
		m, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		returnMode = int(m)
	}
	if returnMode < 0 || returnMode > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	caseInsensitive := false
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		if args[3].Type == ValueError {
			return args[3], nil
		}
		flag, e := coerceRegexCaseFlag(args[3])
		if e != nil {
			return *e, nil
		}
		caseInsensitive = flag
	}
	re, errVal := compileRegex(ValueToString(args[1]), caseInsensitive)
	if errVal != nil {
		return *errVal, nil
	}

	// Array lifting is handled inside the function because modes 1 and 2
	// produce arrays per input cell and would nest if routed through the
	// registry's element-wise broadcaster.
	if args[0].Type == ValueArray {
		return regexExtractArray(re, args[0], returnMode)
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return regexExtractScalar(re, ValueToString(args[0]), returnMode), nil
}

func regexExtractScalar(re *regexp.Regexp, text string, mode int) Value {
	switch mode {
	case 0:
		loc := re.FindStringIndex(text)
		if loc == nil {
			return ErrorVal(ErrValNA)
		}
		return StringVal(text[loc[0]:loc[1]])
	case 1:
		matches := re.FindAllString(text, -1)
		if len(matches) == 0 {
			return ErrorVal(ErrValNA)
		}
		rows := make([][]Value, len(matches))
		for i, m := range matches {
			rows[i] = []Value{StringVal(m)}
		}
		return Value{Type: ValueArray, Array: rows}
	case 2:
		sub := re.FindStringSubmatch(text)
		if sub == nil {
			return ErrorVal(ErrValNA)
		}
		if len(sub) == 1 {
			// No capture groups: return the whole match as a 1x1 row so
			// the shape is consistent with the capture-group path and with
			// regexExtractArray mode 2.
			return Value{Type: ValueArray, Array: [][]Value{{StringVal(sub[0])}}}
		}
		groups := sub[1:]
		row := make([]Value, len(groups))
		for i, g := range groups {
			row[i] = StringVal(g)
		}
		return Value{Type: ValueArray, Array: [][]Value{row}}
	}
	return ErrorVal(ErrValVALUE)
}

// regexExtractArray lifts REGEXEXTRACT over an array input. Mode 0 applies
// element-wise and preserves the input array shape (rows×cols). Mode 1
// returns an N×M array where M is the maximum match count across inputs;
// short rows are padded with #N/A. Mode 2 returns an N×G array where G is
// the regexp capture-group count (re.NumSubexp(), or 1 when the pattern has
// no capture groups); missing groups and non-matching rows are #N/A.
func regexExtractArray(re *regexp.Regexp, input Value, mode int) (Value, error) {
	rows, cols := effectiveArrayBounds(input)
	// Flatten in row-major order to a single list of inputs.
	type cell struct {
		row, col int
		text     string
		isErr    bool
		errVal   Value
	}
	cells := make([]cell, 0, rows*cols)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			v := arrayElementDirect(input, rows, cols, i, j)
			if v.Type == ValueError {
				cells = append(cells, cell{row: i, col: j, isErr: true, errVal: v})
				continue
			}
			cells = append(cells, cell{row: i, col: j, text: ValueToString(v)})
		}
	}

	switch mode {
	case 0:
		out := make([][]Value, rows)
		idx := 0
		for i := 0; i < rows; i++ {
			out[i] = make([]Value, cols)
			for j := 0; j < cols; j++ {
				c := cells[idx]
				idx++
				if c.isErr {
					out[i][j] = c.errVal
					continue
				}
				out[i][j] = regexExtractScalar(re, c.text, 0)
			}
		}
		return Value{Type: ValueArray, Array: out}, nil
	case 1:
		// Only defined for single-column / single-row inputs in Excel; for
		// higher-dimensional inputs we still produce a rectangular result
		// per row, taking max matches across the whole input.
		maxMatches := 1
		rowMatches := make([][]string, len(cells))
		for k, c := range cells {
			if c.isErr {
				continue
			}
			m := re.FindAllString(c.text, -1)
			rowMatches[k] = m
			if len(m) > maxMatches {
				maxMatches = len(m)
			}
		}
		// Collapse to rows × maxMatches when input is a column; otherwise
		// produce one output row per input cell flattened row-major.
		if cols == 1 {
			out := make([][]Value, rows)
			for i := 0; i < rows; i++ {
				out[i] = make([]Value, maxMatches)
				c := cells[i]
				if c.isErr {
					for j := 0; j < maxMatches; j++ {
						out[i][j] = c.errVal
					}
					continue
				}
				ms := rowMatches[i]
				if len(ms) == 0 {
					for j := 0; j < maxMatches; j++ {
						out[i][j] = ErrorVal(ErrValNA)
					}
					continue
				}
				for j := 0; j < maxMatches; j++ {
					if j < len(ms) {
						out[i][j] = StringVal(ms[j])
					} else {
						out[i][j] = ErrorVal(ErrValNA)
					}
				}
			}
			return Value{Type: ValueArray, Array: out}, nil
		}
		// Multi-column input: flatten into rows of maxMatches.
		out := make([][]Value, len(cells))
		for k, c := range cells {
			out[k] = make([]Value, maxMatches)
			if c.isErr {
				for j := 0; j < maxMatches; j++ {
					out[k][j] = c.errVal
				}
				continue
			}
			ms := rowMatches[k]
			if len(ms) == 0 {
				for j := 0; j < maxMatches; j++ {
					out[k][j] = ErrorVal(ErrValNA)
				}
				continue
			}
			for j := 0; j < maxMatches; j++ {
				if j < len(ms) {
					out[k][j] = StringVal(ms[j])
				} else {
					out[k][j] = ErrorVal(ErrValNA)
				}
			}
		}
		return Value{Type: ValueArray, Array: out}, nil
	case 2:
		// Determine group count from the regex itself (NumSubexp).
		numGroups := re.NumSubexp()
		if numGroups == 0 {
			numGroups = 1 // whole match
		}
		out := make([][]Value, len(cells))
		for k, c := range cells {
			row := make([]Value, numGroups)
			if c.isErr {
				for j := 0; j < numGroups; j++ {
					row[j] = c.errVal
				}
				out[k] = row
				continue
			}
			sub := re.FindStringSubmatch(c.text)
			if sub == nil {
				for j := 0; j < numGroups; j++ {
					row[j] = ErrorVal(ErrValNA)
				}
				out[k] = row
				continue
			}
			if re.NumSubexp() == 0 {
				row[0] = StringVal(sub[0])
			} else {
				for j := 0; j < numGroups; j++ {
					row[j] = StringVal(sub[j+1])
				}
			}
			out[k] = row
		}
		return Value{Type: ValueArray, Array: out}, nil
	}
	return ErrorVal(ErrValVALUE), nil
}

func fnREGEXREPLACE(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	if args[1].Type == ValueError {
		return args[1], nil
	}
	if args[2].Type == ValueError {
		return args[2], nil
	}
	instanceNum := 0
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		if args[3].Type == ValueError {
			return args[3], nil
		}
		n, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		instanceNum = int(n)
	}
	caseInsensitive := false
	if len(args) == 5 && args[4].Type != ValueEmpty {
		if args[4].Type == ValueError {
			return args[4], nil
		}
		flag, e := coerceRegexCaseFlag(args[4])
		if e != nil {
			return *e, nil
		}
		caseInsensitive = flag
	}
	re, errVal := compileRegex(ValueToString(args[1]), caseInsensitive)
	if errVal != nil {
		return *errVal, nil
	}
	text := ValueToString(args[0])
	replacement := ValueToString(args[2])

	if instanceNum == 0 {
		// Replace all matches.
		return StringVal(regexReplaceAll(re, text, replacement)), nil
	}
	return StringVal(regexReplaceNth(re, text, replacement, instanceNum)), nil
}

// regexReplaceAll replaces every match in text using Excel-style $N / $0
// backreferences. Non-digit-following '$' is emitted literally.
func regexReplaceAll(re *regexp.Regexp, text, replacement string) string {
	matches := re.FindAllStringSubmatchIndex(text, -1)
	if matches == nil {
		return text
	}
	return applyRegexReplacements(text, replacement, matches)
}

// regexReplaceNth replaces a single match identified by instanceNum:
//
//	positive N: the Nth match (1-based).
//	negative N: the Nth match from the end (-1 = last).
//
// If the index is out of range, the text is returned unchanged.
func regexReplaceNth(re *regexp.Regexp, text, replacement string, instanceNum int) string {
	matches := re.FindAllStringSubmatchIndex(text, -1)
	if len(matches) == 0 {
		return text
	}
	var idx int
	if instanceNum > 0 {
		idx = instanceNum - 1
	} else {
		idx = len(matches) + instanceNum
	}
	if idx < 0 || idx >= len(matches) {
		return text
	}
	return applyRegexReplacements(text, replacement, matches[idx:idx+1])
}

// applyRegexReplacements rebuilds text by replacing the given match ranges
// with the expanded replacement template. Each match is a FindAllStringSubmatchIndex
// entry: [start, end, group1start, group1end, ...].
func applyRegexReplacements(text, replacement string, matches [][]int) string {
	var b strings.Builder
	b.Grow(len(text))
	cursor := 0
	for _, m := range matches {
		start, end := m[0], m[1]
		if start < cursor {
			// Zero-width match produced by an earlier iteration; skip.
			continue
		}
		b.WriteString(text[cursor:start])
		expandExcelReplacement(&b, replacement, text, m)
		cursor = end
		if start == end {
			// Zero-width match: emit one literal rune so we advance,
			// matching how ReplaceAllString handles empty matches. Decode
			// the rune so multi-byte UTF-8 characters aren't split.
			if cursor < len(text) {
				r, size := utf8.DecodeRuneInString(text[cursor:])
				b.WriteRune(r)
				cursor += size
			}
		}
	}
	if cursor < len(text) {
		b.WriteString(text[cursor:])
	}
	return b.String()
}

// expandExcelReplacement appends replacement to b, substituting $N / $0 with
// the corresponding capture group from match. Unlike Go's Expand, a bare '$'
// is emitted literally and '$$' collapses to a single '$'. Multi-digit group
// numbers are consumed greedily (up to the highest defined group).
func expandExcelReplacement(b *strings.Builder, replacement, text string, match []int) {
	i := 0
	maxGroup := (len(match) / 2) - 1
	for i < len(replacement) {
		c := replacement[i]
		if c != '$' {
			b.WriteByte(c)
			i++
			continue
		}
		if i+1 >= len(replacement) {
			b.WriteByte('$')
			i++
			continue
		}
		next := replacement[i+1]
		if next == '$' {
			b.WriteByte('$')
			i += 2
			continue
		}
		if next < '0' || next > '9' {
			b.WriteByte('$')
			i++
			continue
		}
		// Consume digits greedily, backing off if the number exceeds the
		// available groups (Excel prefers the longest valid group index).
		j := i + 1
		for j < len(replacement) && replacement[j] >= '0' && replacement[j] <= '9' {
			j++
		}
		chosen := -1
		for k := j; k > i+1; k-- {
			n := 0
			for p := i + 1; p < k; p++ {
				n = n*10 + int(replacement[p]-'0')
			}
			if n <= maxGroup {
				chosen = n
				j = k
				break
			}
		}
		if chosen < 0 {
			// Number too large even for one digit: emit '$' literally and
			// consume nothing beyond the dollar sign so the digits are kept.
			b.WriteByte('$')
			i++
			continue
		}
		gStart, gEnd := match[2*chosen], match[2*chosen+1]
		if gStart >= 0 && gEnd >= 0 {
			b.WriteString(text[gStart:gEnd])
		}
		i = j
	}
}
