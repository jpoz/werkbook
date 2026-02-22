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

func fnCOUNTIF(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
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
