package formula

import (
	"math"
	"sort"
	"strconv"
	"strings"
)

func fnSUM(args []Value) (Value, error) {
	sum := 0.0
	if e := iterateNumeric(args, func(n float64) { sum += n }); e != nil {
		return *e, nil
	}
	return NumberVal(sum), nil
}

func fnAVEDEV(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Compute mean
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	// Compute average of absolute deviations
	absDevSum := 0.0
	for _, v := range nums {
		d := v - mean
		if d < 0 {
			d = -d
		}
		absDevSum += d
	}
	return NumberVal(absDevSum / float64(n)), nil
}

func fnDEVSQ(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Compute mean
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	// Compute sum of squared deviations
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(ssq), nil
}

func fnAVERAGE(args []Value) (Value, error) {
	sum := 0.0
	count := 0
	if e := iterateNumeric(args, func(n float64) { sum += n; count++ }); e != nil {
		return *e, nil
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnCOUNT(args []Value) (Value, error) {
	count := 0
	for _, arg := range args {
		switch arg.Type {
		case ValueArray:
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueNumber {
						count++
					}
				}
			}
		case ValueNumber:
			count++
		case ValueString:
			if _, err := strconv.ParseFloat(arg.Str, 64); err == nil {
				count++
			}
		}
	}
	return NumberVal(float64(count)), nil
}

func fnCOUNTA(args []Value) (Value, error) {
	count := 0
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type != ValueEmpty {
						count++
					}
				}
			}
		} else if arg.Type != ValueEmpty {
			count++
		}
	}
	return NumberVal(float64(count)), nil
}

func fnMIN(args []Value) (Value, error) {
	min := math.MaxFloat64
	found := false
	if e := iterateNumeric(args, func(n float64) {
		if !found || n < min {
			min = n
			found = true
		}
	}); e != nil {
		return *e, nil
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(min), nil
}

func fnMAX(args []Value) (Value, error) {
	max := -math.MaxFloat64
	found := false
	if e := iterateNumeric(args, func(n float64) {
		if !found || n > max {
			max = n
			found = true
		}
	}); e != nil {
		return *e, nil
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(max), nil
}

// collectNumeric gathers all numeric values from args into a slice.
func collectNumeric(args []Value) ([]float64, *Value) {
	var nums []float64
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					if cell.Type == ValueNumber {
						nums = append(nums, cell.Num)
					}
				}
			}
		} else {
			if arg.Type == ValueError {
				return nil, &arg
			}
			n, e := coerceNum(arg)
			if e != nil {
				return nil, e
			}
			nums = append(nums, n)
		}
	}
	return nums, nil
}

func fnMEDIAN(args []Value) (Value, error) {
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	n := len(nums)
	if n%2 == 1 {
		return NumberVal(nums[n/2]), nil
	}
	return NumberVal((nums[n/2-1] + nums[n/2]) / 2), nil
}

func fnLARGE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	var nums []float64
	arr := args[0]
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			for _, cell := range row {
				if cell.Type == ValueError {
					return cell, nil
				}
				if cell.Type == ValueNumber {
					nums = append(nums, cell.Num)
				}
			}
		}
	} else {
		n, e := coerceNum(arr)
		if e != nil {
			return *e, nil
		}
		nums = append(nums, n)
	}
	k, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ki := int(k)
	if ki < 1 || ki > len(nums) {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	return NumberVal(nums[len(nums)-ki]), nil
}

func fnSMALL(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	var nums []float64
	arr := args[0]
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			for _, cell := range row {
				if cell.Type == ValueError {
					return cell, nil
				}
				if cell.Type == ValueNumber {
					nums = append(nums, cell.Num)
				}
			}
		}
	} else {
		n, e := coerceNum(arr)
		if e != nil {
			return *e, nil
		}
		nums = append(nums, n)
	}
	k, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ki := int(k)
	if ki < 1 || ki > len(nums) {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	return NumberVal(nums[ki-1]), nil
}

func fnCOUNTBLANK(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	count := 0
	arg := args[0]
	switch arg.Type {
	case ValueArray:
		for _, row := range arg.Array {
			for _, cell := range row {
				if cell.Type == ValueEmpty {
					count++
				}
			}
		}
	case ValueEmpty:
		count = 1
	}
	return NumberVal(float64(count)), nil
}

// matchesCriteria checks if a value matches Excel-style criteria.
func matchesCriteria(v Value, criteria Value) bool {
	critStr := valueToString(criteria)

	if len(critStr) >= 2 {
		switch {
		case strings.HasPrefix(critStr, ">="):
			return compareToCriteria(v, critStr[2:]) >= 0
		case strings.HasPrefix(critStr, "<="):
			return compareToCriteria(v, critStr[2:]) <= 0
		case strings.HasPrefix(critStr, "<>"):
			return compareToCriteria(v, critStr[2:]) != 0
		case strings.HasPrefix(critStr, ">"):
			return compareToCriteria(v, critStr[1:]) > 0
		case strings.HasPrefix(critStr, "<"):
			return compareToCriteria(v, critStr[1:]) < 0
		case strings.HasPrefix(critStr, "="):
			return compareToCriteria(v, critStr[1:]) == 0
		}
	}

	if strings.ContainsAny(critStr, "*?") {
		return wildcardMatch(valueToString(v), critStr)
	}

	if criteria.Type == ValueNumber {
		n, e := coerceNum(v)
		if e != nil {
			return false
		}
		return n == criteria.Num
	}
	return strings.EqualFold(valueToString(v), critStr)
}

func compareToCriteria(v Value, critValStr string) int {
	if n, e := coerceNum(v); e == nil {
		if cn, err := strconv.ParseFloat(critValStr, 64); err == nil {
			return cmpFloat(n, cn)
		}
	}
	return strings.Compare(strings.ToLower(valueToString(v)), strings.ToLower(critValStr))
}

func wildcardMatch(s, pattern string) bool {
	return wildcardHelper(strings.ToLower(s), strings.ToLower(pattern))
}

func wildcardHelper(s, p string) bool {
	for len(p) > 0 {
		switch p[0] {
		case '*':
			for len(p) > 0 && p[0] == '*' {
				p = p[1:]
			}
			if len(p) == 0 {
				return true
			}
			for i := 0; i <= len(s); i++ {
				if wildcardHelper(s[i:], p) {
					return true
				}
			}
			return false
		case '?':
			if len(s) == 0 {
				return false
			}
			s = s[1:]
			p = p[1:]
		default:
			if len(s) == 0 || s[0] != p[0] {
				return false
			}
			s = s[1:]
			p = p[1:]
		}
	}
	return len(s) == 0
}

func fnSUMIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	if rangeArg.Type == ValueError {
		return rangeArg, nil
	}
	criteria := args[1]
	sumRange := rangeArg
	if len(args) == 3 {
		sumRange = args[2]
	}

	if rangeArg.Type != ValueArray {
		if matchesCriteria(rangeArg, criteria) {
			n, e := coerceNum(sumRange)
			if e != nil {
				return NumberVal(0), nil
			}
			return NumberVal(n), nil
		}
		return NumberVal(0), nil
	}

	sum := 0.0
	for i, row := range rangeArg.Array {
		for j, cell := range row {
			if matchesCriteria(cell, criteria) {
				var sv Value
				if sumRange.Type == ValueArray && i < len(sumRange.Array) && j < len(sumRange.Array[i]) {
					sv = sumRange.Array[i][j]
				} else if sumRange.Type != ValueArray {
					sv = sumRange
				}
				if n, e := coerceNum(sv); e == nil {
					sum += n
				}
			}
		}
	}
	return NumberVal(sum), nil
}

func fnSUMIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	sumRange := args[0]
	if sumRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	sum := 0.0
	for r, row := range sumRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !matchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := coerceNum(sumRange.Array[r][c]); e == nil {
					sum += n
				}
			}
		}
	}
	return NumberVal(sum), nil
}

func fnMAXIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	maxRange := args[0]
	if maxRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	maxVal := -math.MaxFloat64
	found := false
	for r, row := range maxRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !matchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := coerceNum(maxRange.Array[r][c]); e == nil {
					if !found || n > maxVal {
						maxVal = n
						found = true
					}
				}
			}
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(maxVal), nil
}

func fnMINIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	minRange := args[0]
	if minRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	minVal := math.MaxFloat64
	found := false
	for r, row := range minRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !matchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := coerceNum(minRange.Array[r][c]); e == nil {
					if !found || n < minVal {
						minVal = n
						found = true
					}
				}
			}
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(minVal), nil
}

func fnCOUNTIF(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	if rangeArg.Type == ValueError {
		return rangeArg, nil
	}
	criteria := args[1]
	count := 0
	if rangeArg.Type == ValueArray {
		for _, row := range rangeArg.Array {
			for _, cell := range row {
				if matchesCriteria(cell, criteria) {
					count++
				}
			}
		}
	} else if matchesCriteria(rangeArg, criteria) {
		count = 1
	}
	return NumberVal(float64(count)), nil
}

func fnCOUNTIFS(args []Value) (Value, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	firstRange := args[0]
	if firstRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	count := 0
	for r, row := range firstRange.Array {
		for c := range row {
			allMatch := true
			for k := 0; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !matchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				count++
			}
		}
	}
	return NumberVal(float64(count)), nil
}

func fnAVERAGEIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	criteria := args[1]
	avgRange := rangeArg
	if len(args) == 3 {
		avgRange = args[2]
	}

	sum := 0.0
	count := 0
	if rangeArg.Type == ValueArray {
		for i, row := range rangeArg.Array {
			for j, cell := range row {
				if matchesCriteria(cell, criteria) {
					var sv Value
					if avgRange.Type == ValueArray && i < len(avgRange.Array) && j < len(avgRange.Array[i]) {
						sv = avgRange.Array[i][j]
					}
					if n, e := coerceNum(sv); e == nil {
						sum += n
						count++
					}
				}
			}
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnAVERAGEIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	avgRange := args[0]
	if avgRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	sum := 0.0
	count := 0
	for r, row := range avgRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !matchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := coerceNum(avgRange.Array[r][c]); e == nil {
					sum += n
					count++
				}
			}
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnSUMSQ(args []Value) (Value, error) {
	sum := 0.0
	if e := iterateNumeric(args, func(n float64) { sum += n * n }); e != nil {
		return *e, nil
	}
	return NumberVal(sum), nil
}

func fnSTDEV(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	// Compute mean
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	// Compute sum of squared deviations
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(math.Sqrt(ssq / float64(n-1))), nil
}

func fnSTDEVP(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(math.Sqrt(ssq / float64(n))), nil // divide by n, not n-1
}

func fnVAR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(ssq / float64(n-1)), nil
}

func fnVARP(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(ssq / float64(n)), nil // divide by n, not n-1
}

func fnMODE(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNA), nil
	}
	// Count frequency of each value
	freq := make(map[float64]int)
	// Track order of first appearance
	order := make([]float64, 0)
	for _, n := range nums {
		if freq[n] == 0 {
			order = append(order, n)
		}
		freq[n]++
	}
	// Find value with highest frequency (must be >= 2)
	bestVal := 0.0
	bestCount := 1 // minimum must be 2 to qualify
	for _, v := range order {
		if freq[v] > bestCount {
			bestCount = freq[v]
			bestVal = v
		}
	}
	if bestCount < 2 {
		return ErrorVal(ErrValNA), nil // no duplicates
	}
	return NumberVal(bestVal), nil
}

func fnPERCENTILE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Collect numeric values from first arg (the array)
	nums, e := collectNumeric(args[:1])
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Get k value
	k, e2 := coerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	if k < 0 || k > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	// Sort ascending
	sort.Float64s(nums)
	n := len(nums)
	if n == 1 {
		return NumberVal(nums[0]), nil
	}
	rank := k * float64(n-1)
	intPart := int(rank)
	frac := rank - float64(intPart)
	if intPart >= n-1 {
		return NumberVal(nums[n-1]), nil
	}
	result := nums[intPart] + frac*(nums[intPart+1]-nums[intPart])
	return NumberVal(result), nil
}

func fnRANK(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	num, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	// Collect numeric values from the reference (arg[1])
	nums, e2 := collectNumeric(args[1:2])
	if e2 != nil {
		return *e2, nil
	}
	// Determine order
	ascending := false
	if len(args) == 3 {
		order, e3 := coerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		ascending = order != 0
	}
	// Check if number exists in the list
	found := false
	for _, v := range nums {
		if v == num {
			found = true
			break
		}
	}
	if !found {
		return ErrorVal(ErrValNA), nil
	}
	// Compute rank
	rank := 1
	for _, v := range nums {
		if ascending {
			if v < num {
				rank++
			}
		} else {
			if v > num {
				rank++
			}
		}
	}
	return NumberVal(float64(rank)), nil
}

func fnGEOMEAN(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// All values must be > 0.
	sumLn := 0.0
	for _, v := range nums {
		if v <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
		sumLn += math.Log(v)
	}
	return NumberVal(math.Exp(sumLn / float64(n))), nil
}

func fnSUMPRODUCT(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}
	firstArr := args[0].Array
	rows := len(firstArr)
	cols := 0
	if rows > 0 {
		cols = len(firstArr[0])
	}

	for _, arg := range args[1:] {
		if arg.Type != ValueArray {
			return ErrorVal(ErrValVALUE), nil
		}
		if len(arg.Array) != rows {
			return ErrorVal(ErrValVALUE), nil
		}
		for _, row := range arg.Array {
			if len(row) != cols {
				return ErrorVal(ErrValVALUE), nil
			}
		}
	}

	sum := 0.0
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			product := 1.0
			for _, arg := range args {
				cell := arg.Array[r][c]
				if cell.Type == ValueError {
					return cell, nil
				}
				n, e := coerceNum(cell)
				if e != nil {
					n = 0
				}
				product *= n
			}
			sum += product
		}
	}
	return NumberVal(sum), nil
}
