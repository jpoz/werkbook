package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode/utf8"
)

// CellResolver abstracts cell/range lookups so the VM has no dependency on Sheet.
type CellResolver interface {
	GetCellValue(addr CellAddr) Value
	GetRangeValues(addr RangeAddr) [][]Value
}

// Eval executes a compiled formula and returns the result.
func Eval(cf *CompiledFormula, resolver CellResolver) (Value, error) {
	stack := make([]Value, 0, 16)

	push := func(v Value) { stack = append(stack, v) }
	pop := func() (Value, error) {
		if len(stack) == 0 {
			return Value{}, fmt.Errorf("stack underflow")
		}
		v := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		return v, nil
	}

	for _, inst := range cf.Code {
		switch inst.Op {
		case OpPushNum:
			push(cf.Consts[inst.Operand])
		case OpPushStr:
			push(cf.Consts[inst.Operand])
		case OpPushBool:
			push(BoolVal(inst.Operand != 0))
		case OpPushError:
			push(ErrorVal(ErrorValue(inst.Operand)))
		case OpPushEmpty:
			push(EmptyVal())

		case OpLoadCell:
			addr := cf.Refs[inst.Operand]
			push(resolver.GetCellValue(addr))

		case OpLoadRange:
			addr := cf.Ranges[inst.Operand]
			rows := resolver.GetRangeValues(addr)
			push(Value{Type: ValueArray, Array: rows})

		case OpAdd:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			bn, be := coerceNum(b)
			if ae != nil {
				push(*ae)
			} else if be != nil {
				push(*be)
			} else {
				push(NumberVal(an + bn))
			}

		case OpSub:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			bn, be := coerceNum(b)
			if ae != nil {
				push(*ae)
			} else if be != nil {
				push(*be)
			} else {
				push(NumberVal(an - bn))
			}

		case OpMul:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			bn, be := coerceNum(b)
			if ae != nil {
				push(*ae)
			} else if be != nil {
				push(*be)
			} else {
				push(NumberVal(an * bn))
			}

		case OpDiv:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			bn, be := coerceNum(b)
			if ae != nil {
				push(*ae)
			} else if be != nil {
				push(*be)
			} else if bn == 0 {
				push(ErrorVal(ErrValDIV0))
			} else {
				push(NumberVal(an / bn))
			}

		case OpPow:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			bn, be := coerceNum(b)
			if ae != nil {
				push(*ae)
			} else if be != nil {
				push(*be)
			} else {
				push(NumberVal(math.Pow(an, bn)))
			}

		case OpNeg:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			if ae != nil {
				push(*ae)
			} else {
				push(NumberVal(-an))
			}

		case OpPercent:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := coerceNum(a)
			if ae != nil {
				push(*ae)
			} else {
				push(NumberVal(an / 100))
			}

		case OpConcat:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(StringVal(valueToString(a) + valueToString(b)))

		case OpEq:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) == 0))

		case OpNe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) != 0))

		case OpLt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) < 0))

		case OpLe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) <= 0))

		case OpGt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) > 0))

		case OpGe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(compareValues(a, b) >= 0))

		case OpCall:
			funcID := int(inst.Operand >> 8)
			argc := int(inst.Operand & 0xFF)
			if argc > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in function call")
			}
			args := make([]Value, argc)
			copy(args, stack[len(stack)-argc:])
			stack = stack[:len(stack)-argc]

			result, err := callFunction(funcID, args)
			if err != nil {
				return Value{}, err
			}
			push(result)

		case OpMakeArray:
			rows := int(inst.Operand >> 16)
			cols := int(inst.Operand & 0xFFFF)
			total := rows * cols
			if total > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in array construction")
			}
			elems := make([]Value, total)
			copy(elems, stack[len(stack)-total:])
			stack = stack[:len(stack)-total]

			arr := make([][]Value, rows)
			for r := 0; r < rows; r++ {
				arr[r] = elems[r*cols : (r+1)*cols]
			}
			push(Value{Type: ValueArray, Array: arr})

		default:
			return Value{}, fmt.Errorf("unknown opcode %d", inst.Op)
		}
	}

	if len(stack) != 1 {
		return Value{}, fmt.Errorf("expected 1 value on stack, got %d", len(stack))
	}
	return stack[0], nil
}

// coerceNum converts a Value to float64 for arithmetic.
// Returns the number and nil on success, or 0 and a pointer to an error Value.
func coerceNum(v Value) (float64, *Value) {
	switch v.Type {
	case ValueNumber:
		return v.Num, nil
	case ValueEmpty:
		return 0, nil
	case ValueBool:
		if v.Bool {
			return 1, nil
		}
		return 0, nil
	case ValueString:
		if v.Str == "" {
			return 0, nil
		}
		n, err := strconv.ParseFloat(v.Str, 64)
		if err != nil {
			e := ErrorVal(ErrValVALUE)
			return 0, &e
		}
		return n, nil
	case ValueError:
		return 0, &v
	default:
		e := ErrorVal(ErrValVALUE)
		return 0, &e
	}
}

func valueToString(v Value) string {
	switch v.Type {
	case ValueNumber:
		return strconv.FormatFloat(v.Num, 'f', -1, 64)
	case ValueString:
		return v.Str
	case ValueBool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case ValueError:
		return errorValueToString(v.Err)
	default:
		return ""
	}
}

func errorValueToString(e ErrorValue) string {
	switch e {
	case ErrValDIV0:
		return "#DIV/0!"
	case ErrValNA:
		return "#N/A"
	case ErrValNAME:
		return "#NAME?"
	case ErrValNULL:
		return "#NULL!"
	case ErrValNUM:
		return "#NUM!"
	case ErrValREF:
		return "#REF!"
	case ErrValVALUE:
		return "#VALUE!"
	case ErrValSPILL:
		return "#SPILL!"
	case ErrValCALC:
		return "#CALC!"
	case ErrValGETTINGDATA:
		return "#GETTING_DATA"
	default:
		return "#VALUE!"
	}
}

// compareValues compares two values for ordering. Returns -1, 0, or 1.
// Excel ordering: errors < empty/numbers < strings < booleans (approximately).
// For simplicity: compare same-type values naturally, different types by type rank.
func compareValues(a, b Value) int {
	// Coerce empty to number 0 for comparison
	if a.Type == ValueEmpty {
		a = NumberVal(0)
	}
	if b.Type == ValueEmpty {
		b = NumberVal(0)
	}

	if a.Type == b.Type {
		switch a.Type {
		case ValueNumber:
			return cmpFloat(a.Num, b.Num)
		case ValueString:
			return strings.Compare(strings.ToLower(a.Str), strings.ToLower(b.Str))
		case ValueBool:
			if a.Bool == b.Bool {
				return 0
			}
			if !a.Bool {
				return -1
			}
			return 1
		}
	}

	// Different types: numbers < strings < booleans
	return typeRank(a.Type) - typeRank(b.Type)
}

func typeRank(t ValueType) int {
	switch t {
	case ValueError:
		return 0
	case ValueNumber, ValueEmpty:
		return 1
	case ValueString:
		return 2
	case ValueBool:
		return 3
	default:
		return 4
	}
}

func cmpFloat(a, b float64) int {
	if a < b {
		return -1
	}
	if a > b {
		return 1
	}
	return 0
}

func isTruthy(v Value) bool {
	switch v.Type {
	case ValueBool:
		return v.Bool
	case ValueNumber:
		return v.Num != 0
	case ValueString:
		return v.Str != ""
	default:
		return false
	}
}

// callFunction dispatches a function call by function ID.
func callFunction(funcID int, args []Value) (Value, error) {
	if funcID < 0 || funcID >= len(knownFunctions) {
		return Value{}, fmt.Errorf("unknown function ID %d", funcID)
	}
	name := knownFunctions[funcID]

	switch name {
	case "SUM":
		return fnSUM(args)
	case "AVERAGE":
		return fnAVERAGE(args)
	case "IF":
		return fnIF(args)
	case "COUNT":
		return fnCOUNT(args)
	case "COUNTA":
		return fnCOUNTA(args)
	case "MIN":
		return fnMIN(args)
	case "MAX":
		return fnMAX(args)
	case "AND":
		return fnAND(args)
	case "OR":
		return fnOR(args)
	case "NOT":
		return fnNOT(args)
	case "ABS":
		return fnABS(args)
	case "INT":
		return fnINT(args)
	case "MOD":
		return fnMOD(args)
	case "ROUND":
		return fnROUND(args)
	case "CONCAT", "CONCATENATE":
		return fnCONCATENATE(args)
	case "IFERROR":
		return fnIFERROR(args)
	case "ISBLANK":
		return fnISBLANK(args)
	case "ISERR", "ISERROR":
		return fnISERROR(args)
	case "ISNUMBER":
		return fnISNUMBER(args)
	case "ISTEXT":
		return fnISTEXT(args)
	case "LEN":
		return fnLEN(args)
	case "LEFT":
		return fnLEFT(args)
	case "RIGHT":
		return fnRIGHT(args)
	case "MID":
		return fnMID(args)
	case "UPPER":
		return fnUPPER(args)
	case "LOWER":
		return fnLOWER(args)
	case "TRIM":
		return fnTRIM(args)
	default:
		return Value{}, fmt.Errorf("unimplemented function: %s", name)
	}
}

// iterateNumeric calls fn for each numeric value in args, expanding arrays.
// Non-numeric values in ranges are skipped; non-numeric scalar args cause #VALUE!.
func iterateNumeric(args []Value, fn func(float64)) *Value {
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return &cell
					}
					if cell.Type == ValueNumber {
						fn(cell.Num)
					}
					// skip text/empty/bool in ranges
				}
			}
		} else {
			if arg.Type == ValueError {
				return &arg
			}
			n, e := coerceNum(arg)
			if e != nil {
				return e
			}
			fn(n)
		}
	}
	return nil
}

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

func fnIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	if isTruthy(args[0]) {
		return args[1], nil
	}
	if len(args) == 3 {
		return args[2], nil
	}
	return BoolVal(false), nil
}

func fnCOUNT(args []Value) (Value, error) {
	count := 0
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueNumber {
						count++
					}
				}
			}
		} else if arg.Type == ValueNumber {
			count++
		} else if arg.Type == ValueString {
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

func fnAND(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if !isTruthy(arg) {
			return BoolVal(false), nil
		}
	}
	return BoolVal(true), nil
}

func fnOR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if isTruthy(arg) {
			return BoolVal(true), nil
		}
	}
	return BoolVal(false), nil
}

func fnNOT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return BoolVal(!isTruthy(args[0])), nil
}

func fnABS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Abs(n)), nil
}

func fnINT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Floor(n)), nil
}

func fnMOD(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	d, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if d == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	// Excel MOD: result has the sign of the divisor
	result := n - d*math.Floor(n/d)
	return NumberVal(result), nil
}

func fnROUND(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pow := math.Pow(10, math.Floor(digits))
	return NumberVal(math.Round(n*pow) / pow), nil
}

func fnCONCATENATE(args []Value) (Value, error) {
	var b strings.Builder
	for _, arg := range args {
		b.WriteString(valueToString(arg))
	}
	return StringVal(b.String()), nil
}

func fnIFERROR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[1], nil
	}
	return args[0], nil
}

func fnISBLANK(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueEmpty), nil
}

func fnISERROR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueError), nil
}

func fnISNUMBER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueNumber), nil
}

func fnISTEXT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueString), nil
}

func fnLEN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	return NumberVal(float64(utf8.RuneCountInString(s))), nil
}

func fnLEFT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		n = int(num)
	}
	runes := []rune(s)
	if n > len(runes) {
		n = len(runes)
	}
	if n < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(runes[:n])), nil
}

func fnRIGHT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		n = int(num)
	}
	runes := []rune(s)
	if n > len(runes) {
		n = len(runes)
	}
	if n < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(runes[len(runes)-n:])), nil
}

func fnMID(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	startNum, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	numChars, e := coerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	start := int(startNum) - 1 // Excel MID is 1-based
	length := int(numChars)
	if start < 0 || length < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	runes := []rune(s)
	if start >= len(runes) {
		return StringVal(""), nil
	}
	end := start + length
	if end > len(runes) {
		end = len(runes)
	}
	return StringVal(string(runes[start:end])), nil
}

func fnUPPER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.ToUpper(valueToString(args[0]))), nil
}

func fnLOWER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.ToLower(valueToString(args[0]))), nil
}

func fnTRIM(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	// Excel TRIM removes leading/trailing spaces and collapses internal runs to single space
	fields := strings.Fields(s)
	return StringVal(strings.Join(fields, " ")), nil
}
