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
	ValueRef              // cell reference; Num encodes col+row*100_000
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

// String returns the display string for an ErrorValue (e.g., "#DIV/0!").
func (e ErrorValue) String() string {
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
	Type        ValueType
	Num         float64
	Str         string
	Bool        bool
	Err         ErrorValue
	Array       [][]Value  // used by ValueArray for range results
	RangeOrigin *RangeAddr // set on ValueArray when loaded from a worksheet range
	CellOrigin  *CellAddr  // set when a single-cell result carries its source address (e.g. OFFSET)
	FromCell    bool       // true when loaded from a cell reference (OpLoadCell)
	NoSpill     bool       // true for arrays that represent whole-row/column references (INDEX with 0); in non-array context these become #VALUE!
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

// ErrorValueFromString converts an error display string (e.g. "#NUM!") to an ErrorValue.
// Returns ErrValVALUE if the string is not recognized.
func ErrorValueFromString(s string) ErrorValue {
	switch s {
	case "#DIV/0!":
		return ErrValDIV0
	case "#N/A":
		return ErrValNA
	case "#NAME?":
		return ErrValNAME
	case "#NULL!":
		return ErrValNULL
	case "#NUM!":
		return ErrValNUM
	case "#REF!":
		return ErrValREF
	case "#VALUE!":
		return ErrValVALUE
	case "#SPILL!":
		return ErrValSPILL
	case "#CALC!":
		return ErrValCALC
	case "#GETTING_DATA":
		return ErrValGETTINGDATA
	default:
		return ErrValVALUE
	}
}

// CellAddr is a compiled cell address.
type CellAddr struct {
	Sheet    string // sheet name (empty if unqualified)
	SheetEnd string // non-empty for 3D references (Sheet2:Sheet5!A1)
	Col      int    // 1-based column
	Row      int    // 1-based row
}

// RangeAddr is a compiled range address.
type RangeAddr struct {
	Sheet    string // sheet name (empty if unqualified)
	SheetEnd string // non-empty for 3D references (Sheet2:Sheet5!A1:B2)
	FromCol  int
	FromRow  int
	ToCol    int
	ToRow    int
}
