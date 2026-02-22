package formula

// ValueType identifies the kind of value held by a Value.
type ValueType byte

const (
	ValueEmpty  ValueType = iota
	ValueNumber           // float64
	ValueString           // string
	ValueBool             // bool
	ValueError            // ErrorValue
	ValueArray            // unused by compiler; reserved for VM
)

// ErrorValue is a numeric error code for the formula engine.
// Named ErrorValue to avoid collision with the existing ErrorCode string type in ast.go.
type ErrorValue byte

const (
	ErrValDIV0        ErrorValue = iota // #DIV/0!
	ErrValNA                            // #N/A
	ErrValNAME                          // #NAME?
	ErrValNULL                          // #NULL!
	ErrValNUM                           // #NUM!
	ErrValREF                           // #REF!
	ErrValVALUE                         // #VALUE!
	ErrValSPILL                         // #SPILL!
	ErrValCALC                          // #CALC!
	ErrValGETTINGDATA                   // #GETTING_DATA
)

// errorCodeFromAST converts the parser's ErrorCode string to a numeric ErrorValue.
func errorCodeFromAST(code ErrorCode) ErrorValue {
	switch code {
	case ErrDIV0:
		return ErrValDIV0
	case ErrNA:
		return ErrValNA
	case ErrNAME:
		return ErrValNAME
	case ErrNULL:
		return ErrValNULL
	case ErrNUM:
		return ErrValNUM
	case ErrREF:
		return ErrValREF
	case ErrVALUE:
		return ErrValVALUE
	case ErrSPILL:
		return ErrValSPILL
	case ErrCALC:
		return ErrValCALC
	case ErrGETTINGDATA:
		return ErrValGETTINGDATA
	default:
		return ErrValVALUE
	}
}

// Value is a tagged union representing a formula engine value.
type Value struct {
	Type ValueType
	Num  float64
	Str  string
	Bool bool
	Err  ErrorValue
}

// NumberVal creates a Value holding a float64.
func NumberVal(f float64) Value {
	return Value{Type: ValueNumber, Num: f}
}

// StringVal creates a Value holding a string.
func StringVal(s string) Value {
	return Value{Type: ValueString, Str: s}
}

// BoolVal creates a Value holding a bool.
func BoolVal(b bool) Value {
	return Value{Type: ValueBool, Bool: b}
}

// ErrorVal creates a Value holding an error code.
func ErrorVal(e ErrorValue) Value {
	return Value{Type: ValueError, Err: e}
}

// EmptyVal creates an empty Value.
func EmptyVal() Value {
	return Value{Type: ValueEmpty}
}

// CellAddr is a compiled cell address.
type CellAddr struct {
	Sheet string // sheet name (empty if unqualified)
	Col   int    // 1-based column
	Row   int    // 1-based row
}

// RangeAddr is a compiled range address.
type RangeAddr struct {
	Sheet   string // sheet name (empty if unqualified)
	FromCol int
	FromRow int
	ToCol   int
	ToRow   int
}
