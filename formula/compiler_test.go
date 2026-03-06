package formula

import (
	"strings"
	"testing"
)

// compileFormula is a test helper that parses and compiles a formula string.
func compileFormula(t *testing.T, input string) *CompiledFormula {
	t.Helper()
	node, err := Parse(input)
	if err != nil {
		t.Fatalf("Parse(%q) error: %v", input, err)
	}
	cf, err := Compile(input, node)
	if err != nil {
		t.Fatalf("Compile(%q) error: %v", input, err)
	}
	return cf
}

func TestCompileLiterals(t *testing.T) {
	tests := []struct {
		input  string
		wantOp OpCode
	}{
		{"42", OpPushNum},
		{"3.14", OpPushNum},
		{`"hello"`, OpPushStr},
		{"TRUE", OpPushBool},
		{"FALSE", OpPushBool},
		{"#N/A", OpPushError},
		{"#DIV/0!", OpPushError},
		{"#VALUE!", OpPushError},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			cf := compileFormula(t, tt.input)
			if len(cf.Code) != 1 {
				t.Fatalf("expected 1 instruction, got %d", len(cf.Code))
			}
			if cf.Code[0].Op != tt.wantOp {
				t.Errorf("opcode = %s, want %s", cf.Code[0].Op, tt.wantOp)
			}
		})
	}
}

func TestCompileNumberConstPool(t *testing.T) {
	cf := compileFormula(t, "42")
	if len(cf.Consts) != 1 {
		t.Fatalf("expected 1 constant, got %d", len(cf.Consts))
	}
	if cf.Consts[0].Type != ValueNumber || cf.Consts[0].Num != 42 {
		t.Errorf("constant = %+v, want NumberVal(42)", cf.Consts[0])
	}
	if cf.Code[0].Operand != 0 {
		t.Errorf("operand = %d, want 0", cf.Code[0].Operand)
	}
}

func TestCompileStringConstPool(t *testing.T) {
	cf := compileFormula(t, `"hello"`)
	if len(cf.Consts) != 1 {
		t.Fatalf("expected 1 constant, got %d", len(cf.Consts))
	}
	if cf.Consts[0].Type != ValueString || cf.Consts[0].Str != "hello" {
		t.Errorf("constant = %+v, want StringVal(hello)", cf.Consts[0])
	}
}

func TestCompileBoolOperand(t *testing.T) {
	cf := compileFormula(t, "TRUE")
	if cf.Code[0].Operand != 1 {
		t.Errorf("TRUE operand = %d, want 1", cf.Code[0].Operand)
	}
	cf = compileFormula(t, "FALSE")
	if cf.Code[0].Operand != 0 {
		t.Errorf("FALSE operand = %d, want 0", cf.Code[0].Operand)
	}
}

func TestCompileErrorOperand(t *testing.T) {
	tests := []struct {
		input   string
		wantErr ErrorValue
	}{
		{"#DIV/0!", ErrValDIV0},
		{"#N/A", ErrValNA},
		{"#NAME?", ErrValNAME},
		{"#NULL!", ErrValNULL},
		{"#NUM!", ErrValNUM},
		{"#REF!", ErrValREF},
		{"#VALUE!", ErrValVALUE},
		{"#SPILL!", ErrValSPILL},
		{"#CALC!", ErrValCALC},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			cf := compileFormula(t, tt.input)
			if cf.Code[0].Operand != uint32(tt.wantErr) {
				t.Errorf("operand = %d, want %d", cf.Code[0].Operand, tt.wantErr)
			}
		})
	}
}

func TestCompileCellRef(t *testing.T) {
	cf := compileFormula(t, "A1")
	if len(cf.Code) != 1 {
		t.Fatalf("expected 1 instruction, got %d", len(cf.Code))
	}
	if cf.Code[0].Op != OpLoadCell {
		t.Errorf("opcode = %s, want LoadCell", cf.Code[0].Op)
	}
	if len(cf.Refs) != 1 {
		t.Fatalf("expected 1 ref, got %d", len(cf.Refs))
	}
	if cf.Refs[0].Col != 1 || cf.Refs[0].Row != 1 || cf.Refs[0].Sheet != "" {
		t.Errorf("ref = %+v, want {Col:1 Row:1}", cf.Refs[0])
	}
}

func TestCompileCellRefWithSheet(t *testing.T) {
	cf := compileFormula(t, "Sheet1!B3")
	if len(cf.Refs) != 1 {
		t.Fatalf("expected 1 ref, got %d", len(cf.Refs))
	}
	ref := cf.Refs[0]
	if ref.Sheet != "Sheet1" || ref.Col != 2 || ref.Row != 3 {
		t.Errorf("ref = %+v, want {Sheet:Sheet1 Col:2 Row:3}", ref)
	}
}

func TestCompileRangeRef(t *testing.T) {
	cf := compileFormula(t, "A1:B5")
	if len(cf.Code) != 1 {
		t.Fatalf("expected 1 instruction, got %d", len(cf.Code))
	}
	if cf.Code[0].Op != OpLoadRange {
		t.Errorf("opcode = %s, want LoadRange", cf.Code[0].Op)
	}
	if len(cf.Ranges) != 1 {
		t.Fatalf("expected 1 range, got %d", len(cf.Ranges))
	}
	rng := cf.Ranges[0]
	if rng.FromCol != 1 || rng.FromRow != 1 || rng.ToCol != 2 || rng.ToRow != 5 || rng.Sheet != "" {
		t.Errorf("range = %+v", rng)
	}
}

func TestCompileRangeRefWithSheet(t *testing.T) {
	cf := compileFormula(t, "Sheet1!A1:B5")
	if len(cf.Ranges) != 1 {
		t.Fatalf("expected 1 range, got %d", len(cf.Ranges))
	}
	rng := cf.Ranges[0]
	if rng.Sheet != "Sheet1" {
		t.Errorf("range sheet = %q, want Sheet1", rng.Sheet)
	}
}

func TestCompileBinaryOps(t *testing.T) {
	tests := []struct {
		input  string
		wantOp OpCode
	}{
		{"1+2", OpAdd},
		{"1-2", OpSub},
		{"1*2", OpMul},
		{"1/2", OpDiv},
		{"1^2", OpPow},
		{`"a"&"b"`, OpConcat},
		{"1=2", OpEq},
		{"1<>2", OpNe},
		{"1<2", OpLt},
		{"1<=2", OpLe},
		{"1>2", OpGt},
		{"1>=2", OpGe},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			cf := compileFormula(t, tt.input)
			// Last instruction should be the binary op.
			last := cf.Code[len(cf.Code)-1]
			if last.Op != tt.wantOp {
				t.Errorf("last opcode = %s, want %s", last.Op, tt.wantOp)
			}
		})
	}
}

func TestCompileUnaryMinus(t *testing.T) {
	cf := compileFormula(t, "-1")
	// Should be: PushNum, Neg
	if len(cf.Code) != 2 {
		t.Fatalf("expected 2 instructions, got %d", len(cf.Code))
	}
	if cf.Code[0].Op != OpPushNum {
		t.Errorf("code[0] = %s, want PushNum", cf.Code[0].Op)
	}
	if cf.Code[1].Op != OpNeg {
		t.Errorf("code[1] = %s, want Neg", cf.Code[1].Op)
	}
}

func TestCompileUnaryPlus(t *testing.T) {
	cf := compileFormula(t, "+1")
	// Unary + is a no-op — only PushNum should be emitted.
	if len(cf.Code) != 1 {
		t.Fatalf("expected 1 instruction, got %d", len(cf.Code))
	}
	if cf.Code[0].Op != OpPushNum {
		t.Errorf("code[0] = %s, want PushNum", cf.Code[0].Op)
	}
}

func TestCompilePostfixPercent(t *testing.T) {
	cf := compileFormula(t, "50%")
	if len(cf.Code) != 2 {
		t.Fatalf("expected 2 instructions, got %d", len(cf.Code))
	}
	if cf.Code[1].Op != OpPercent {
		t.Errorf("code[1] = %s, want Percent", cf.Code[1].Op)
	}
}

func TestCompileFuncCall(t *testing.T) {
	cf := compileFormula(t, "SUM(1,2,3)")
	// PushNum(1), PushNum(2), PushNum(3), Call
	if len(cf.Code) != 4 {
		t.Fatalf("expected 4 instructions, got %d", len(cf.Code))
	}
	call := cf.Code[3]
	if call.Op != OpCall {
		t.Fatalf("last opcode = %s, want Call", call.Op)
	}
	funcID := call.Operand >> 8
	argc := call.Operand & 0xFF
	wantID := uint32(LookupFunc("SUM"))
	if funcID != wantID {
		t.Errorf("funcID = %d, want %d (SUM)", funcID, wantID)
	}
	if argc != 3 {
		t.Errorf("argc = %d, want 3", argc)
	}
}

func TestCompileZeroArgFunc(t *testing.T) {
	cf := compileFormula(t, "NOW()")
	if len(cf.Code) != 1 {
		t.Fatalf("expected 1 instruction, got %d", len(cf.Code))
	}
	call := cf.Code[0]
	if call.Op != OpCall {
		t.Fatalf("opcode = %s, want Call", call.Op)
	}
	argc := call.Operand & 0xFF
	if argc != 0 {
		t.Errorf("argc = %d, want 0", argc)
	}
}

func TestCompileArrayLit(t *testing.T) {
	cf := compileFormula(t, "{1,2;3,4}")
	// PushNum(1), PushNum(2), PushNum(3), PushNum(4), MakeArray
	if len(cf.Code) != 5 {
		t.Fatalf("expected 5 instructions, got %d", len(cf.Code))
	}
	last := cf.Code[4]
	if last.Op != OpMakeArray {
		t.Fatalf("last opcode = %s, want MakeArray", last.Op)
	}
	rows := last.Operand >> 16
	cols := last.Operand & 0xFFFF
	if rows != 2 || cols != 2 {
		t.Errorf("array dims = %dx%d, want 2x2", rows, cols)
	}
}

func TestCompileComplexSUM(t *testing.T) {
	cf := compileFormula(t, "SUM(A1:A10)")
	// LoadRange, Call
	if len(cf.Code) != 2 {
		t.Fatalf("expected 2 instructions, got %d", len(cf.Code))
	}
	if cf.Code[0].Op != OpLoadRange {
		t.Errorf("code[0] = %s, want LoadRange", cf.Code[0].Op)
	}
	if cf.Code[1].Op != OpCall {
		t.Errorf("code[1] = %s, want Call", cf.Code[1].Op)
	}
}

func TestCompileComplexIF(t *testing.T) {
	cf := compileFormula(t, "IF(A1>0,A1,0)")
	// LoadCell(A1), PushNum(0), Gt, LoadCell(A1), PushNum(0), Call
	if cf.Code[len(cf.Code)-1].Op != OpCall {
		t.Errorf("last opcode = %s, want Call", cf.Code[len(cf.Code)-1].Op)
	}
	argc := cf.Code[len(cf.Code)-1].Operand & 0xFF
	if argc != 3 {
		t.Errorf("IF argc = %d, want 3", argc)
	}
}

func TestCompileComplexArithmetic(t *testing.T) {
	cf := compileFormula(t, "A1+B1*2")
	// LoadCell(A1), LoadCell(B1), PushNum(2), Mul, Add
	if len(cf.Code) != 5 {
		t.Fatalf("expected 5 instructions, got %d", len(cf.Code))
	}
	if cf.Code[3].Op != OpMul {
		t.Errorf("code[3] = %s, want Mul", cf.Code[3].Op)
	}
	if cf.Code[4].Op != OpAdd {
		t.Errorf("code[4] = %s, want Add", cf.Code[4].Op)
	}
}

func TestCompileConstDeduplication(t *testing.T) {
	cf := compileFormula(t, "1+1")
	// Same number should produce only 1 constant.
	if len(cf.Consts) != 1 {
		t.Errorf("expected 1 constant (dedup), got %d", len(cf.Consts))
	}
	// Both PushNum instructions should reference index 0.
	if cf.Code[0].Operand != 0 || cf.Code[1].Operand != 0 {
		t.Errorf("operands = %d, %d, want 0, 0", cf.Code[0].Operand, cf.Code[1].Operand)
	}
}

func TestCompileRefDeduplication(t *testing.T) {
	cf := compileFormula(t, "A1+A1")
	if len(cf.Refs) != 1 {
		t.Errorf("expected 1 ref (dedup), got %d", len(cf.Refs))
	}
	if cf.Code[0].Operand != 0 || cf.Code[1].Operand != 0 {
		t.Errorf("operands = %d, %d, want 0, 0", cf.Code[0].Operand, cf.Code[1].Operand)
	}
}

func TestCompileRangeDeduplication(t *testing.T) {
	cf := compileFormula(t, "SUM(A1:B5)+SUM(A1:B5)")
	if len(cf.Ranges) != 1 {
		t.Errorf("expected 1 range (dedup), got %d", len(cf.Ranges))
	}
}

func TestCompileStringDeduplication(t *testing.T) {
	cf := compileFormula(t, `CONCATENATE("x","x")`)
	if len(cf.Consts) != 1 {
		t.Errorf("expected 1 constant (dedup), got %d", len(cf.Consts))
	}
}

func TestCompileUnknownFunction(t *testing.T) {
	node, err := Parse("NOTAFUNCTION(1)")
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}
	_, err = Compile("NOTAFUNCTION(1)", node)
	if err == nil {
		t.Fatal("expected compile error for unknown function")
	}
	if !strings.Contains(err.Error(), "unknown function") {
		t.Errorf("error = %q, want it to contain 'unknown function'", err.Error())
	}
}

func TestCompileXludfPrefix(t *testing.T) {
	// _xludf. prefix means "user-defined function" — must produce #NAME?
	// at runtime (not a compile error) so IFERROR and similar can catch it.
	src := "_xludf.CEILING.MATH(2.5)"
	node, err := Parse(src)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}
	cf, err := Compile(src, node)
	if err != nil {
		t.Fatalf("unexpected compile error: %v", err)
	}
	result, err := Eval(cf, nil, &EvalContext{})
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNAME {
		t.Errorf("got %v, want #NAME? error", result)
	}
}

func TestXludfPrefixWithIFERROR(t *testing.T) {
	// IFERROR should catch the #NAME? from a _xludf. function.
	src := `IFERROR(_xludf.ACOT(0),"caught")`
	node, err := Parse(src)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}
	cf, err := Compile(src, node)
	if err != nil {
		t.Fatalf("unexpected compile error: %v", err)
	}
	result, err := Eval(cf, nil, &EvalContext{})
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if result.Type != ValueString || result.Str != "caught" {
		t.Errorf("got %v, want string 'caught'", result)
	}
}

func TestCompileSourcePreservation(t *testing.T) {
	input := "SUM(A1:A10)"
	cf := compileFormula(t, input)
	if cf.Source != input {
		t.Errorf("Source = %q, want %q", cf.Source, input)
	}
}

func TestCompileDistinctConsts(t *testing.T) {
	cf := compileFormula(t, "1+2")
	if len(cf.Consts) != 2 {
		t.Errorf("expected 2 constants, got %d", len(cf.Consts))
	}
}

func TestCompileMultipleRefs(t *testing.T) {
	cf := compileFormula(t, "A1+B2")
	if len(cf.Refs) != 2 {
		t.Fatalf("expected 2 refs, got %d", len(cf.Refs))
	}
	if cf.Refs[0].Col != 1 || cf.Refs[0].Row != 1 {
		t.Errorf("ref[0] = %+v", cf.Refs[0])
	}
	if cf.Refs[1].Col != 2 || cf.Refs[1].Row != 2 {
		t.Errorf("ref[1] = %+v", cf.Refs[1])
	}
}

func TestCompileInstructionString(t *testing.T) {
	inst := Instruction{Op: OpAdd, Operand: 0}
	s := inst.String()
	if s != "Add 0" {
		t.Errorf("String() = %q, want %q", s, "Add 0")
	}
}

func TestCompileOpCodeString(t *testing.T) {
	if OpPushNum.String() != "PushNum" {
		t.Errorf("OpPushNum.String() = %q", OpPushNum.String())
	}
	if OpCall.String() != "Call" {
		t.Errorf("OpCall.String() = %q", OpCall.String())
	}
}

func TestCompileFuncCaseInsensitive(t *testing.T) {
	// The parser already uppercases function names, but verify the compiler
	// handles it correctly end-to-end.
	cf := compileFormula(t, "sum(1)")
	if cf.Code[len(cf.Code)-1].Op != OpCall {
		t.Errorf("expected Call opcode")
	}
}

func TestCompileNestedFunctions(t *testing.T) {
	cf := compileFormula(t, "IF(AND(A1>0,B1<100),1,0)")
	last := cf.Code[len(cf.Code)-1]
	if last.Op != OpCall {
		t.Fatalf("last opcode = %s, want Call", last.Op)
	}
	// IF has 3 args
	argc := last.Operand & 0xFF
	if argc != 3 {
		t.Errorf("IF argc = %d, want 3", argc)
	}
}
