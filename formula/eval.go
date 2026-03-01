package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

const (
	maxExcelRows = 1048576 // maximum rows in an Excel worksheet
	maxExcelCols = 16384   // maximum columns in an Excel worksheet (XFD)
)

// CellResolver abstracts cell/range lookups so the VM has no dependency on Sheet.
type CellResolver interface {
	GetCellValue(addr CellAddr) Value
	GetRangeValues(addr RangeAddr) [][]Value
}

// EvalContext provides context about the current evaluation environment.
type EvalContext struct {
	CurrentCol     int
	CurrentRow     int
	CurrentSheet   string
	IsArrayFormula bool // true for CSE (Ctrl+Shift+Enter) array formulas
}

// Eval executes a compiled formula and returns the result.
func Eval(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext) (Value, error) {
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
			// Implicit intersection: when a full-column or full-row range is
			// used in a non-array formula, reduce to the single cell at the
			// formula's own row/column rather than loading the entire range.
			if ctx != nil && !ctx.IsArrayFormula {
				isFullCol := addr.FromRow == 1 && addr.ToRow >= maxExcelRows
				isFullRow := addr.FromCol == 1 && addr.ToCol >= maxExcelCols
				if isFullCol && addr.FromCol == addr.ToCol && ctx.CurrentRow >= addr.FromRow {
					// Full-column ref like F:F → intersect at current row
					push(resolver.GetCellValue(CellAddr{
						Sheet: addr.Sheet,
						Col:   addr.FromCol,
						Row:   ctx.CurrentRow,
					}))
					continue
				}
				if isFullRow && addr.FromRow == addr.ToRow && ctx.CurrentCol >= addr.FromCol {
					// Full-row ref like 1:1 → intersect at current column
					push(resolver.GetCellValue(CellAddr{
						Sheet: addr.Sheet,
						Col:   ctx.CurrentCol,
						Row:   addr.FromRow,
					}))
					continue
				}
			}
			rows := resolver.GetRangeValues(addr)
			// Pad trailing blank rows for bounded ranges. GetRangeValues
			// clamps toRow to MaxRow to avoid huge allocations for
			// full-column refs, but bounded ranges like A1:A5 need all
			// requested rows so functions like COUNTBLANK see every blank.
			isFullCol := addr.FromRow == 1 && addr.ToRow >= maxExcelRows
			isFullRow := addr.FromCol == 1 && addr.ToCol >= maxExcelCols
			if !isFullCol && !isFullRow {
				expectedRows := addr.ToRow - addr.FromRow + 1
				cols := addr.ToCol - addr.FromCol + 1
				for len(rows) < expectedRows {
					emptyRow := make([]Value, cols)
					for j := range emptyRow {
						emptyRow[j] = EmptyVal()
					}
					rows = append(rows, emptyRow)
				}
			}
			push(Value{Type: ValueArray, Array: rows})

		case OpLoadCellRef:
			addr := cf.Refs[inst.Operand]
			// Encode col and row into Num: col + row*100_000.
			// Max col = 16384 < 100_000, max row = 1_048_576, product < 2^53.
			push(Value{Type: ValueRef, Num: float64(addr.Col + addr.Row*100_000)})

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
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) == 0))
			}

		case OpNe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) != 0))
			}

		case OpLt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) < 0))
			}

		case OpLe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) <= 0))
			}

		case OpGt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) > 0))
			}

		case OpGe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(BoolVal(compareValues(a, b) >= 0))
			}

		case OpCall:
			funcID := int(inst.Operand >> 8)
			argc := int(inst.Operand & 0xFF)
			if argc > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in function call")
			}
			args := make([]Value, argc)
			copy(args, stack[len(stack)-argc:])
			stack = stack[:len(stack)-argc]

			result, err := callFunction(funcID, args, ctx)
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
func compareValues(a, b Value) int {
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
func callFunction(funcID int, args []Value, ctx *EvalContext) (Value, error) {
	if funcID < 0 || funcID >= len(knownFunctions) {
		return Value{}, fmt.Errorf("unknown function ID %d", funcID)
	}
	name := knownFunctions[funcID]

	switch name {
	// Math
	case "ABS":
		return fnABS(args)
	case "ACOS":
		return fnACOS(args)
	case "ACOSH":
		return fnACOSH(args)
	case "ASIN":
		return fnASIN(args)
	case "ATAN":
		return fnATAN(args)
	case "ATAN2":
		return fnATAN2(args)
	case "CEILING":
		return fnCEILING(args)
	case "COMBIN":
		return fnCOMBIN(args)
	case "COS":
		return fnCOS(args)
	case "DEGREES":
		return fnDEGREES(args)
	case "EVEN":
		return fnEVEN(args)
	case "EXP":
		return fnEXP(args)
	case "FACT":
		return fnFACT(args)
	case "FLOOR":
		return fnFLOOR(args)
	case "GCD":
		return fnGCD(args)
	case "INT":
		return fnINT(args)
	case "LCM":
		return fnLCM(args)
	case "LN":
		return fnLN(args)
	case "LOG":
		return fnLOG(args)
	case "LOG10":
		return fnLOG10(args)
	case "MOD":
		return fnMOD(args)
	case "MROUND":
		return fnMROUND(args)
	case "ODD":
		return fnODD(args)
	case "PERMUT":
		return fnPERMUT(args)
	case "PI":
		return fnPI(args)
	case "POWER":
		return fnPOWER(args)
	case "PRODUCT":
		return fnPRODUCT(args)
	case "QUOTIENT":
		return fnQUOTIENT(args)
	case "RADIANS":
		return fnRADIANS(args)
	case "RAND":
		return fnRAND(args)
	case "RANDBETWEEN":
		return fnRANDBETWEEN(args)
	case "ROUND":
		return fnROUND(args)
	case "ROUNDDOWN":
		return fnROUNDDOWN(args)
	case "ROUNDUP":
		return fnROUNDUP(args)
	case "SIGN":
		return fnSIGN(args)
	case "SIN":
		return fnSIN(args)
	case "SQRT":
		return fnSQRT(args)
	case "SUBTOTAL":
		return fnSUBTOTAL(args)
	case "TAN":
		return fnTAN(args)
	case "TRUNC":
		return fnTRUNC(args)

	// Statistics
	case "AVERAGE":
		return fnAVERAGE(args)
	case "AVERAGEIF":
		return fnAVERAGEIF(args)
	case "AVERAGEIFS":
		return fnAVERAGEIFS(args)
	case "COUNT":
		return fnCOUNT(args)
	case "COUNTA":
		return fnCOUNTA(args)
	case "COUNTBLANK":
		return fnCOUNTBLANK(args)
	case "COUNTIF":
		return fnCOUNTIF(args)
	case "COUNTIFS":
		return fnCOUNTIFS(args)
	case "LARGE":
		return fnLARGE(args)
	case "MAX":
		return fnMAX(args)
	case "MAXIFS":
		return fnMAXIFS(args)
	case "MEDIAN":
		return fnMEDIAN(args)
	case "MODE":
		return fnMODE(args)
	case "PERCENTILE":
		return fnPERCENTILE(args)
	case "MIN":
		return fnMIN(args)
	case "MINIFS":
		return fnMINIFS(args)
	case "RANK":
		return fnRANK(args)
	case "SMALL":
		return fnSMALL(args)
	case "SUM":
		return fnSUM(args)
	case "SUMIF":
		return fnSUMIF(args)
	case "SUMIFS":
		return fnSUMIFS(args)
	case "SUMPRODUCT":
		return fnSUMPRODUCT(args)
	case "STDEV":
		return fnSTDEV(args)
	case "STDEVP":
		return fnSTDEVP(args)
	case "SUMSQ":
		return fnSUMSQ(args)
	case "VAR":
		return fnVAR(args)
	case "VARP":
		return fnVARP(args)

	// Text
	case "CHAR":
		return fnCHAR(args)
	case "CHOOSE":
		return fnCHOOSE(args)
	case "CLEAN":
		return fnCLEAN(args)
	case "CODE":
		return fnCODE(args)
	case "CONCAT", "CONCATENATE":
		return fnCONCATENATE(args)
	case "EXACT":
		return fnEXACT(args)
	case "FIND":
		return fnFIND(args)
	case "FIXED":
		return fnFIXED(args)
	case "LEFT":
		return fnLEFT(args)
	case "LEN":
		return fnLEN(args)
	case "LOWER":
		return fnLOWER(args)
	case "MID":
		return fnMID(args)
	case "NUMBERVALUE":
		return fnNUMBERVALUE(args)
	case "PROPER":
		return fnPROPER(args)
	case "REPLACE":
		return fnREPLACE(args)
	case "REPT":
		return fnREPT(args)
	case "RIGHT":
		return fnRIGHT(args)
	case "SEARCH":
		return fnSEARCH(args)
	case "SUBSTITUTE":
		return fnSUBSTITUTE(args)
	case "T":
		return fnT(args)
	case "TEXT":
		return fnTEXT(args)
	case "TEXTJOIN":
		return fnTEXTJOIN(args)
	case "TRIM":
		return fnTRIM(args)
	case "UPPER":
		return fnUPPER(args)
	case "VALUE":
		return fnVALUEFn(args)

	// Logic
	case "AND":
		return fnAND(args)
	case "IF":
		return fnIF(args)
	case "IFERROR":
		return fnIFERROR(args)
	case "IFS":
		return fnIFS(args)
	case "NOT":
		return fnNOT(args)
	case "OR":
		return fnOR(args)
	case "SORT":
		return fnSORT(args)
	case "SWITCH":
		return fnSWITCH(args)
	case "XOR":
		return fnXOR(args)

	// Info
	case "COLUMN":
		return fnCOLUMN(args, ctx)
	case "COLUMNS":
		return fnCOLUMNS(args)
	case "ERROR.TYPE":
		return fnERRORTYPE(args)
	case "IFNA":
		return fnIFNA(args)
	case "ISBLANK":
		return fnISBLANK(args)
	case "ISERR":
		return fnISERR(args)
	case "ISERROR":
		return fnISERROR(args)
	case "ISEVEN":
		return fnISEVEN(args)
	case "ISLOGICAL":
		return fnISLOGICAL(args)
	case "ISNA":
		return fnISNA(args)
	case "ISNONTEXT":
		return fnISNONTEXT(args)
	case "ISODD":
		return fnISODD(args)
	case "ISNUMBER":
		return fnISNUMBER(args)
	case "ISTEXT":
		return fnISTEXT(args)
	case "N":
		return fnN(args)
	case "NA":
		return fnNA(args)
	case "ROW":
		return fnROW(args, ctx)
	case "ROWS":
		return fnROWS(args)
	case "TYPE":
		return fnTYPE(args)

	// Date/Time
	case "DATE":
		return fnDATE(args)
	case "DATEDIF":
		return fnDATEDIF(args)
	case "DATEVALUE":
		return fnDATEVALUE(args)
	case "DAY":
		return fnDAY(args)
	case "DAYS":
		return fnDAYS(args)
	case "EDATE":
		return fnEDATE(args)
	case "EOMONTH":
		return fnEOMONTH(args)
	case "HOUR":
		return fnHOUR(args)
	case "ISOWEEKNUM":
		return fnISOWEEKNUM(args)
	case "MINUTE":
		return fnMINUTE(args)
	case "MONTH":
		return fnMONTH(args)
	case "NETWORKDAYS":
		return fnNETWORKDAYS(args)
	case "NOW":
		return fnNOW(args)
	case "SECOND":
		return fnSECOND(args)
	case "TIME":
		return fnTIME(args)
	case "TODAY":
		return fnTODAY(args)
	case "WEEKDAY":
		return fnWEEKDAY(args)
	case "WEEKNUM":
		return fnWEEKNUM(args)
	case "WORKDAY":
		return fnWORKDAY(args)
	case "YEAR":
		return fnYEAR(args)
	case "YEARFRAC":
		return fnYEARFRAC(args)

	// Lookup
	case "ADDRESS":
		return fnADDRESS(args)
	case "HLOOKUP":
		return fnHLOOKUP(args)
	case "INDEX":
		return fnINDEX(args)
	case "INDIRECT":
		// TODO: requires dynamic reference resolution
		return ErrorVal(ErrValREF), nil
	case "LOOKUP":
		return fnLOOKUP(args)
	case "MATCH":
		return fnMATCH(args)
	case "VLOOKUP":
		return fnVLOOKUP(args)
	case "XLOOKUP":
		return fnXLOOKUP(args)

	default:
		return Value{}, fmt.Errorf("unimplemented function: %s", name)
	}
}

// liftUnary applies a scalar function element-wise over a ValueArray,
// returning a new ValueArray of the same shape. Used for array-formula
// evaluation of functions like ABS, ISNUMBER, etc.
func liftUnary(arr Value, fn func(Value) Value) Value {
	rows := make([][]Value, len(arr.Array))
	for i, row := range arr.Array {
		out := make([]Value, len(row))
		for j, cell := range row {
			out[j] = fn(cell)
		}
		rows[i] = out
	}
	return Value{Type: ValueArray, Array: rows}
}

// arrayElement returns element [i][j] from arr if it is an array,
// or returns the scalar arr otherwise. Used for broadcasting scalars
// alongside arrays in element-wise operations.
func arrayElement(v Value, i, j int) Value {
	if v.Type != ValueArray {
		return v
	}
	if i < len(v.Array) && j < len(v.Array[i]) {
		return v.Array[i][j]
	}
	return ErrorVal(ErrValNA)
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
