package formula

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

func init() {
	Register("ADDRESS", NoCtx(fnADDRESS))
	Register("HLOOKUP", NoCtx(fnHLOOKUP))
	Register("INDEX", NoCtx(fnINDEX))
	Register("INDIRECT", fnINDIRECT)
	Register("LOOKUP", NoCtx(fnLOOKUP))
	Register("MATCH", NoCtx(fnMATCH))
	Register("VLOOKUP", NoCtx(fnVLOOKUP))
	Register("TRANSPOSE", NoCtx(fnTRANSPOSE))
	Register("XLOOKUP", NoCtx(fnXLOOKUP))
}

func fnVLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type == ValueError {
		return table, nil
	}
	if table.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}
	colIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ci := int(colIdx)
	if ci < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	if rangeLookup {
		lastMatch := -1
		for i, row := range table.Array {
			if len(row) == 0 {
				continue
			}
			cmp := CompareValues(row[0], lookup)
			if cmp == 0 {
				lastMatch = i
				break
			}
			if cmp > 0 {
				break
			}
			lastMatch = i
		}
		if lastMatch < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if ci > len(table.Array[lastMatch]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[lastMatch][ci-1], nil
	}

	for _, row := range table.Array {
		if len(row) == 0 {
			continue
		}
		if CompareValuesExact(row[0], lookup) == 0 {
			if ci > len(row) {
				return ErrorVal(ErrValREF), nil
			}
			return row[ci-1], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnHLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type != ValueArray || len(table.Array) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ri := int(rowIdx)
	if ri < 1 || ri > len(table.Array) {
		return ErrorVal(ErrValREF), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	firstRow := table.Array[0]

	if rangeLookup {
		lastMatch := -1
		for i, cell := range firstRow {
			cmp := CompareValues(cell, lookup)
			if cmp == 0 {
				lastMatch = i
				break
			}
			if cmp > 0 {
				break
			}
			lastMatch = i
		}
		if lastMatch < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if lastMatch >= len(table.Array[ri-1]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[ri-1][lastMatch], nil
	}

	for i, cell := range firstRow {
		if CompareValuesExact(cell, lookup) == 0 {
			if i >= len(table.Array[ri-1]) {
				return ErrorVal(ErrValREF), nil
			}
			return table.Array[ri-1][i], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnINDEX(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	arr := args[0]
	if arr.Type != ValueArray {
		return arr, nil
	}
	rowNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ri := int(rowNum)

	// Default col_num: if not provided, default to 1 (first column).
	colNum := 1
	if len(args) == 3 {
		cn, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		colNum = int(cn)
	}

	// Negative indices are invalid and return #VALUE! in Excel.
	if ri < 0 || colNum < 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// row_num=0 means return the entire column (or array if col_num=0 too).
	// The result is an array; in a single-cell (non-array) context the
	// caller will reduce this to #VALUE! automatically.
	if ri == 0 && colNum == 0 {
		return arr, nil
	}
	if ri == 0 {
		// Return entire column as a single-column array.
		ci := colNum - 1
		var col [][]Value
		for _, row := range arr.Array {
			if ci < 0 || ci >= len(row) {
				return ErrorVal(ErrValREF), nil
			}
			col = append(col, []Value{row[ci]})
		}
		return Value{Type: ValueArray, Array: col}, nil
	}
	if colNum == 0 {
		// Return entire row as a single-row array.
		ri--
		if ri < 0 || ri >= len(arr.Array) {
			return ErrorVal(ErrValREF), nil
		}
		return Value{Type: ValueArray, Array: [][]Value{arr.Array[ri]}}, nil
	}

	ri--
	colNum--
	if ri >= len(arr.Array) {
		return ErrorVal(ErrValREF), nil
	}
	if colNum >= len(arr.Array[ri]) {
		return ErrorVal(ErrValREF), nil
	}
	return arr.Array[ri][colNum], nil
}

func fnMATCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	arr := args[1]
	matchType := 1
	if len(args) == 3 {
		mt, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		matchType = int(mt)
	}

	var values []Value
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			values = append(values, row...)
		}
	} else {
		values = []Value{arr}
	}

	switch matchType {
	case 0:
		for i, v := range values {
			if CompareValuesExact(v, lookup) == 0 {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValNA), nil

	case 1:
		// Approximate match (ascending). Empty cells are skipped so that
		// whole-column references (e.g. Q:Q) with sparse data work
		// correctly. We scan all non-empty values and track the last
		// position where value <= lookup, without breaking early. This
		// handles unsorted data gracefully (common in real-world
		// spreadsheets that omit the match_type argument) while producing
		// identical results for properly sorted data.
		last := -1
		for i, v := range values {
			if v.Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(v, lookup)
			if cmp <= 0 {
				last = i
			}
		}
		if last < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(last + 1)), nil

	case -1:
		// Approximate match (descending). Excel uses a binary search
		// expecting descending-sorted data. We replicate that binary
		// search so that unsorted data produces the same result (often
		// #N/A) as Excel.
		n := len(values)
		if n == 0 {
			return ErrorVal(ErrValNA), nil
		}
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			v := values[mid]
			if v.Type == ValueEmpty {
				// Treat empty as less than any lookup value, search left.
				hi = mid - 1
				continue
			}
			cmp := CompareValues(v, lookup)
			if cmp == 0 {
				// Exact match – smallest value >= lookup.
				return NumberVal(float64(mid + 1)), nil
			} else if cmp > 0 {
				// v > lookup – this is a candidate (>= lookup), but
				// there might be a smaller one further right (descending).
				result = mid
				lo = mid + 1
			} else {
				// v < lookup – need larger values, go left (descending).
				hi = mid - 1
			}
		}
		if result < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(result + 1)), nil
	}

	return ErrorVal(ErrValVALUE), nil
}

func fnADDRESS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowNum, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	colNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	row := int(rowNum)
	col := int(colNum)
	if row < 1 || col < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	absNum := 1
	if len(args) >= 3 {
		a, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		absNum = int(a)
	}

	a1Style := true
	if len(args) >= 4 {
		a1Style = IsTruthy(args[3])
	}

	sheetText := ""
	if len(args) >= 5 {
		sheetText = ValueToString(args[4])
	}

	var result string
	if a1Style {
		colName := colNumberToLetters(col)
		switch absNum {
		case 1:
			result = fmt.Sprintf("$%s$%d", colName, row)
		case 2:
			result = fmt.Sprintf("%s$%d", colName, row)
		case 3:
			result = fmt.Sprintf("$%s%d", colName, row)
		case 4:
			result = fmt.Sprintf("%s%d", colName, row)
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	} else {
		switch absNum {
		case 1:
			result = fmt.Sprintf("R%dC%d", row, col)
		case 2:
			result = fmt.Sprintf("R%dC[%d]", row, col)
		case 3:
			result = fmt.Sprintf("R[%d]C%d", row, col)
		case 4:
			result = fmt.Sprintf("R[%d]C[%d]", row, col)
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	}

	if sheetText != "" {
		needsQuote := strings.ContainsAny(sheetText, " '[")
		if needsQuote {
			result = "'" + sheetText + "'!" + result
		} else {
			result = sheetText + "!" + result
		}
	}

	return StringVal(result), nil
}

func fnLOOKUP(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	lookupArr := args[1]
	resultArr := lookupArr
	if len(args) == 3 {
		resultArr = args[2]
	}

	var lookupValues []Value
	if lookupArr.Type == ValueArray {
		for _, row := range lookupArr.Array {
			lookupValues = append(lookupValues, row...)
		}
	} else {
		lookupValues = []Value{lookupArr}
	}

	var resultValues []Value
	if resultArr.Type == ValueArray {
		for _, row := range resultArr.Array {
			resultValues = append(resultValues, row...)
		}
	} else {
		resultValues = []Value{resultArr}
	}

	lastMatch := -1
	for i, v := range lookupValues {
		cmp := CompareValues(v, lookup)
		if cmp <= 0 {
			lastMatch = i
		}
		if cmp > 0 {
			break
		}
	}

	if lastMatch < 0 || lastMatch >= len(resultValues) {
		return ErrorVal(ErrValNA), nil
	}
	return resultValues[lastMatch], nil
}

func fnXLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	lookupArr := args[1]
	returnArr := args[2]

	notFound := ErrorVal(ErrValNA)
	if len(args) >= 4 {
		notFound = args[3]
	}

	matchMode := 0
	if len(args) >= 5 {
		mm, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		matchMode = int(mm)
	}

	var lookupValues []Value
	if lookupArr.Type == ValueArray {
		for _, row := range lookupArr.Array {
			lookupValues = append(lookupValues, row...)
		}
	} else {
		lookupValues = []Value{lookupArr}
	}

	var returnValues []Value
	if returnArr.Type == ValueArray {
		for _, row := range returnArr.Array {
			returnValues = append(returnValues, row...)
		}
	} else {
		returnValues = []Value{returnArr}
	}

	switch matchMode {
	case 0:
		for i, v := range lookupValues {
			if CompareValuesExact(v, lookup) == 0 {
				if i < len(returnValues) {
					return returnValues[i], nil
				}
				return ErrorVal(ErrValNA), nil
			}
		}

	case -1:
		lastMatch := -1
		for i, v := range lookupValues {
			if CompareValues(v, lookup) <= 0 {
				lastMatch = i
			}
		}
		if lastMatch >= 0 && lastMatch < len(returnValues) {
			return returnValues[lastMatch], nil
		}

	case 1:
		for i, v := range lookupValues {
			if CompareValues(v, lookup) >= 0 {
				if i < len(returnValues) {
					return returnValues[i], nil
				}
				break
			}
		}

	case 2:
		// Wildcard match: * matches any sequence, ? matches single char, ~ escapes.
		pattern := ValueToString(lookup)
		re, err := excelPatternToRegexp(pattern)
		if err != nil {
			return ErrorVal(ErrValVALUE), nil
		}
		// Anchor the pattern for full-string matching.
		anchored, err := regexp.Compile("(?i)^" + re.String() + "$")
		if err != nil {
			return ErrorVal(ErrValVALUE), nil
		}
		for i, v := range lookupValues {
			s := ValueToString(v)
			if anchored.MatchString(s) {
				if i < len(returnValues) {
					return returnValues[i], nil
				}
				return ErrorVal(ErrValNA), nil
			}
		}
	}

	return notFound, nil
}

// fnINDIRECT implements INDIRECT(ref_text, [a1]).
// It converts a text string into a cell or range reference and resolves it.
func fnINDIRECT(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}

	refText := ValueToString(args[0])
	if refText == "" {
		return ErrorVal(ErrValREF), nil
	}

	// a1 parameter: default true (A1 style). When false, use R1C1 style.
	a1Style := true
	if len(args) == 2 {
		a1Style = IsTruthy(args[1])
	}

	if ctx == nil || ctx.Resolver == nil {
		return ErrorVal(ErrValREF), nil
	}

	// If R1C1 style, convert to A1 style before parsing.
	if !a1Style {
		converted, err := r1c1ToA1(refText)
		if err != nil {
			return ErrorVal(ErrValREF), nil
		}
		refText = converted
	}

	// Strip dollar signs (absolute markers) for parsing.
	cleaned := strings.ReplaceAll(refText, "$", "")

	// Extract optional sheet prefix.
	sheet := ""
	cellPart := cleaned
	if idx := strings.LastIndex(cleaned, "!"); idx >= 0 {
		sheetPart := cleaned[:idx]
		// Remove surrounding quotes from sheet name.
		if len(sheetPart) >= 2 && sheetPart[0] == '\'' && sheetPart[len(sheetPart)-1] == '\'' {
			sheetPart = sheetPart[1 : len(sheetPart)-1]
		}
		sheet = sheetPart
		cellPart = cleaned[idx+1:]
	}

	// Check if it's a range (contains colon).
	if colonIdx := strings.IndexByte(cellPart, ':'); colonIdx >= 0 {
		left := cellPart[:colonIdx]
		right := cellPart[colonIdx+1:]
		addr, err := indirectParseRange(left, right, sheet)
		if err != nil {
			return ErrorVal(ErrValREF), nil
		}
		isFullCol := addr.FromRow == 1 && addr.ToRow >= maxExcelRows
		isFullRow := addr.FromCol == 1 && addr.ToCol >= maxExcelCols
		// For full-row or full-column ranges (e.g. "1:20", "A:C"), return
		// only the RangeOrigin metadata without resolving cell values.
		// Functions like ROW() and COLUMN() only need the metadata, and
		// resolving all cells in such large ranges causes false circular
		// reference errors when the calling cell falls within the range.
		if isFullCol || isFullRow {
			nRows := addr.ToRow - addr.FromRow + 1
			nCols := addr.ToCol - addr.FromCol + 1
			if isFullRow {
				nCols = 1 // placeholder; actual columns determined by consumer
			}
			if isFullCol {
				nRows = 1 // placeholder; actual rows determined by consumer
			}
			rows := make([][]Value, nRows)
			for i := range rows {
				rows[i] = make([]Value, nCols)
				for j := range rows[i] {
					rows[i][j] = EmptyVal()
				}
			}
			return Value{Type: ValueArray, Array: rows, RangeOrigin: &addr}, nil
		}
		rows := ctx.Resolver.GetRangeValues(addr)
		// Pad trailing blank rows for bounded ranges.
		expectedRows := addr.ToRow - addr.FromRow + 1
		cols := addr.ToCol - addr.FromCol + 1
		for len(rows) < expectedRows {
			emptyRow := make([]Value, cols)
			for j := range emptyRow {
				emptyRow[j] = EmptyVal()
			}
			rows = append(rows, emptyRow)
		}
		return Value{Type: ValueArray, Array: rows, RangeOrigin: &addr}, nil
	}

	// Single cell reference.
	col, row, err := indirectParseCell(cellPart)
	if err != nil {
		return ErrorVal(ErrValREF), nil
	}
	addr := CellAddr{Sheet: sheet, Col: col, Row: row}
	return ctx.Resolver.GetCellValue(addr), nil
}

// indirectParseCell parses a cell reference like "A1" or "B3" into (col, row).
func indirectParseCell(s string) (col, row int, err error) {
	if s == "" {
		return 0, 0, fmt.Errorf("empty cell reference")
	}
	i := 0
	for i < len(s) && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == 0 || i == len(s) {
		return 0, 0, fmt.Errorf("invalid cell reference %q", s)
	}
	col = colLettersToNumber(s[:i])
	if col < 1 || col > maxExcelCols {
		return 0, 0, fmt.Errorf("column out of range in %q", s)
	}
	row, err = strconv.Atoi(s[i:])
	if err != nil || row < 1 || row > maxExcelRows {
		return 0, 0, fmt.Errorf("invalid row in %q", s)
	}
	return col, row, nil
}

// indirectParseRange parses the left and right parts of a range reference.
// Supports cell:cell (A1:C5), row:row (1:20), and col:col (A:C).
func indirectParseRange(left, right, sheet string) (RangeAddr, error) {
	leftIsRowOnly := isAllDigits(left)
	rightIsRowOnly := isAllDigits(right)
	leftIsColOnly := isAllLetters(left)
	rightIsColOnly := isAllLetters(right)

	// Row-only range like "1:20"
	if leftIsRowOnly && rightIsRowOnly {
		r1, err := strconv.Atoi(left)
		if err != nil || r1 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid row %q", left)
		}
		r2, err := strconv.Atoi(right)
		if err != nil || r2 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid row %q", right)
		}
		if r1 > r2 {
			r1, r2 = r2, r1
		}
		return RangeAddr{
			Sheet:   sheet,
			FromCol: 1, FromRow: r1,
			ToCol: maxExcelCols, ToRow: r2,
		}, nil
	}

	// Column-only range like "A:C"
	if leftIsColOnly && rightIsColOnly {
		c1 := colLettersToNumber(left)
		c2 := colLettersToNumber(right)
		if c1 < 1 || c2 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid column range")
		}
		if c1 > c2 {
			c1, c2 = c2, c1
		}
		return RangeAddr{
			Sheet:   sheet,
			FromCol: c1, FromRow: 1,
			ToCol: c2, ToRow: maxExcelRows,
		}, nil
	}

	// Standard cell:cell range like "A1:C5"
	c1, r1, err := indirectParseCell(left)
	if err != nil {
		return RangeAddr{}, err
	}
	c2, r2, err := indirectParseCell(right)
	if err != nil {
		return RangeAddr{}, err
	}
	if c1 > c2 {
		c1, c2 = c2, c1
	}
	if r1 > r2 {
		r1, r2 = r2, r1
	}
	return RangeAddr{
		Sheet:   sheet,
		FromCol: c1, FromRow: r1,
		ToCol: c2, ToRow: r2,
	}, nil
}

// parseR1C1Cell parses a single R1C1-style cell reference like "R1C1" or "R5C3"
// and returns (col, row). The input is case-insensitive.
func parseR1C1Cell(s string) (col, row int, err error) {
	s = strings.ToUpper(s)
	if len(s) < 4 || s[0] != 'R' {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q", s)
	}
	cIdx := strings.IndexByte(s[1:], 'C')
	if cIdx < 0 {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q: missing C", s)
	}
	cIdx++ // adjust for the slice offset
	rowStr := s[1:cIdx]
	colStr := s[cIdx+1:]
	if rowStr == "" || colStr == "" {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q: empty row or col", s)
	}
	row, err = strconv.Atoi(rowStr)
	if err != nil || row < 1 || row > maxExcelRows {
		return 0, 0, fmt.Errorf("invalid row in R1C1 reference %q", s)
	}
	col, err = strconv.Atoi(colStr)
	if err != nil || col < 1 || col > maxExcelCols {
		return 0, 0, fmt.Errorf("invalid col in R1C1 reference %q", s)
	}
	return col, row, nil
}

// r1c1ToA1 converts an R1C1-style reference string to A1-style.
// Supports single cell (R1C1), ranges (R1C1:R5C3), and optional sheet prefixes.
func r1c1ToA1(ref string) (string, error) {
	// Preserve sheet prefix.
	prefix := ""
	cellPart := ref
	if idx := strings.LastIndex(ref, "!"); idx >= 0 {
		prefix = ref[:idx+1]
		cellPart = ref[idx+1:]
	}

	// Check if it's a range.
	if colonIdx := strings.IndexByte(cellPart, ':'); colonIdx >= 0 {
		left := cellPart[:colonIdx]
		right := cellPart[colonIdx+1:]
		c1, r1, err := parseR1C1Cell(left)
		if err != nil {
			return "", err
		}
		c2, r2, err := parseR1C1Cell(right)
		if err != nil {
			return "", err
		}
		return prefix + colNumberToLetters(c1) + strconv.Itoa(r1) + ":" + colNumberToLetters(c2) + strconv.Itoa(r2), nil
	}

	// Single cell.
	c, r, err := parseR1C1Cell(cellPart)
	if err != nil {
		return "", err
	}
	return prefix + colNumberToLetters(c) + strconv.Itoa(r), nil
}

func isAllDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

func isAllLetters(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
			return false
		}
	}
	return true
}

func fnTRANSPOSE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	v := args[0]
	if v.Type != ValueArray {
		return v, nil
	}

	rows := len(v.Array)
	if rows == 0 {
		return Value{Type: ValueArray, Array: nil}, nil
	}

	// Find the maximum column count (handle ragged arrays).
	cols := 0
	for _, row := range v.Array {
		if len(row) > cols {
			cols = len(row)
		}
	}
	if cols == 0 {
		return Value{Type: ValueArray, Array: nil}, nil
	}

	// Transpose: result has cols rows and rows columns.
	result := make([][]Value, cols)
	for c := 0; c < cols; c++ {
		result[c] = make([]Value, rows)
		for r := 0; r < rows; r++ {
			if c < len(v.Array[r]) {
				result[c][r] = v.Array[r][c]
			} else {
				result[c][r] = EmptyVal()
			}
		}
	}
	return Value{Type: ValueArray, Array: result}, nil
}
