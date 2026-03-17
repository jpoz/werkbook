package formula

import "fmt"

// OpCode identifies a bytecode instruction.
type OpCode byte

const (
	OpPushNum         OpCode = iota // operand: index into Consts
	OpPushStr                       // operand: index into Consts
	OpPushBool                      // operand: 0=false, 1=true
	OpPushError                     // operand: ErrorValue
	OpPushEmpty                     // operand: unused
	OpLoadCell                      // operand: index into Refs
	OpLoadRange                     // operand: index into Ranges
	OpAdd                           // pop two, push sum
	OpSub                           // pop two, push difference
	OpMul                           // pop two, push product
	OpDiv                           // pop two, push quotient
	OpPow                           // pop two, push power
	OpNeg                           // pop one, push negation
	OpPercent                       // pop one, push value/100
	OpConcat                        // pop two, push concatenation
	OpEq                            // pop two, push equality
	OpNe                            // pop two, push not-equal
	OpLt                            // pop two, push less-than
	OpLe                            // pop two, push less-or-equal
	OpGt                            // pop two, push greater-than
	OpGe                            // pop two, push greater-or-equal
	OpCall                          // operand: funcID<<8 | argc
	OpMakeArray                     // operand: rows<<16 | cols
	OpLoadCellRef                   // operand: index into Refs; pushes ValueRef (no cell lookup)
	OpEnterArrayCtx                 // operand: unused; pushes array context (suppresses implicit intersection)
	OpLeaveArrayCtx                 // operand: unused; pops array context
	OpLoad3DRange                   // operand: index into Ranges; loads values across multiple sheets
	OpRefResultToBool               // operand: unused; pops value, pushes TRUE if non-error (for ISREF wrapping ref-returning funcs)
	OpLoadParam                     // operand: param slot index; push the bound parameter value
	OpMap                           // operand: subFormulaIdx<<8 | numArrays; execute lambda body per element
	OpReduce                        // operand: subFormulaIdx; pop array and initial value, fold with lambda body
	OpScan                          // operand: subFormulaIdx; pop array and initial value, scan with lambda body returning array
	OpByRow                         // operand: subFormulaIdx; pop array, apply lambda to each row, return column vector
	OpByCol                         // operand: subFormulaIdx; pop array, apply lambda to each column, return row vector
	OpMakeArrayLambda               // operand: subFormulaIdx; pop rows and cols, build array by calling lambda(row, col)
)

var opNames = [...]string{
	OpPushNum:         "PushNum",
	OpPushStr:         "PushStr",
	OpPushBool:        "PushBool",
	OpPushError:       "PushError",
	OpPushEmpty:       "PushEmpty",
	OpLoadCell:        "LoadCell",
	OpLoadRange:       "LoadRange",
	OpAdd:             "Add",
	OpSub:             "Sub",
	OpMul:             "Mul",
	OpDiv:             "Div",
	OpPow:             "Pow",
	OpNeg:             "Neg",
	OpPercent:         "Percent",
	OpConcat:          "Concat",
	OpEq:              "Eq",
	OpNe:              "Ne",
	OpLt:              "Lt",
	OpLe:              "Le",
	OpGt:              "Gt",
	OpGe:              "Ge",
	OpCall:            "Call",
	OpMakeArray:       "MakeArray",
	OpLoadCellRef:     "LoadCellRef",
	OpEnterArrayCtx:   "EnterArrayCtx",
	OpLeaveArrayCtx:   "LeaveArrayCtx",
	OpLoad3DRange:     "Load3DRange",
	OpRefResultToBool: "RefResultToBool",
	OpLoadParam:       "LoadParam",
	OpMap:             "Map",
	OpReduce:          "Reduce",
	OpScan:            "Scan",
	OpByRow:           "ByRow",
	OpByCol:           "ByCol",
	OpMakeArrayLambda: "MakeArrayLambda",
}

func (op OpCode) String() string {
	if int(op) < len(opNames) && opNames[op] != "" {
		return opNames[op]
	}
	return fmt.Sprintf("Op(%d)", op)
}

// Instruction is a single bytecode instruction.
type Instruction struct {
	Op      OpCode
	Operand uint32
}

func (inst Instruction) String() string {
	return fmt.Sprintf("%s %d", inst.Op, inst.Operand)
}

// CompiledFormula is the output of the compiler: bytecode ready for the VM.
type CompiledFormula struct {
	Source      string             // original formula text
	Code        []Instruction      // bytecode instructions
	Consts      []Value            // constant pool (numbers and strings)
	Refs        []CellAddr         // cell reference table
	Ranges      []RangeAddr        // range reference table
	SubFormulas []*CompiledFormula // lambda bodies for MAP/REDUCE/SCAN/BYROW/BYCOL
}
