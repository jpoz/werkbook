package werkbook

import (
	"fmt"
	"math"
	"time"
)

// ValueType represents the type of a cell value.
type ValueType int

const (
	TypeEmpty ValueType = iota
	TypeNumber
	TypeString
	TypeBool
	TypeError
)

// Value is a tagged union representing a cell value.
type Value struct {
	Type   ValueType
	Number float64
	String string
	Bool   bool
}

// IsEmpty returns true if the value is empty.
func (v Value) IsEmpty() bool {
	return v.Type == TypeEmpty
}

// Raw returns the underlying Go value.
func (v Value) Raw() any {
	switch v.Type {
	case TypeNumber:
		return v.Number
	case TypeString:
		return v.String
	case TypeBool:
		return v.Bool
	default:
		return nil
	}
}

func (v Value) GoString() string {
	switch v.Type {
	case TypeEmpty:
		return "Value{Empty}"
	case TypeNumber:
		return fmt.Sprintf("Value{Number: %g}", v.Number)
	case TypeString:
		return fmt.Sprintf("Value{String: %q}", v.String)
	case TypeBool:
		return fmt.Sprintf("Value{Bool: %t}", v.Bool)
	case TypeError:
		return fmt.Sprintf("Value{Error: %q}", v.String)
	default:
		return "Value{?}"
	}
}

// toValue converts an arbitrary Go value to a Value.
func toValue(v any) (Value, error) {
	switch val := v.(type) {
	case nil:
		return Value{Type: TypeEmpty}, nil
	case string:
		return Value{Type: TypeString, String: val}, nil
	case bool:
		return Value{Type: TypeBool, Bool: val}, nil
	case int:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case int8:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case int16:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case int32:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case int64:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case uint:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case uint8:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case uint16:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case uint32:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case uint64:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case float32:
		return Value{Type: TypeNumber, Number: float64(val)}, nil
	case float64:
		if math.IsNaN(val) || math.IsInf(val, 0) {
			return Value{}, fmt.Errorf("%w: %v", ErrUnsupportedType, val)
		}
		return Value{Type: TypeNumber, Number: val}, nil
	case time.Time:
		serial := timeToExcelSerial(val)
		return Value{Type: TypeNumber, Number: serial}, nil
	case Value:
		return val, nil
	default:
		return Value{}, fmt.Errorf("%w: %T", ErrUnsupportedType, v)
	}
}
