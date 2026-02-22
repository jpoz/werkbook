package werkbook

import (
	"math"
	"testing"
)

func TestToValue(t *testing.T) {
	tests := []struct {
		name    string
		input   any
		want    Value
		wantErr bool
	}{
		{"nil", nil, Value{Type: TypeEmpty}, false},
		{"string", "hello", Value{Type: TypeString, String: "hello"}, false},
		{"empty string", "", Value{Type: TypeString, String: ""}, false},
		{"bool true", true, Value{Type: TypeBool, Bool: true}, false},
		{"bool false", false, Value{Type: TypeBool, Bool: false}, false},
		{"int", 42, Value{Type: TypeNumber, Number: 42}, false},
		{"int64", int64(100), Value{Type: TypeNumber, Number: 100}, false},
		{"float64", 3.14, Value{Type: TypeNumber, Number: 3.14}, false},
		{"float32", float32(1.5), Value{Type: TypeNumber, Number: 1.5}, false},
		{"uint", uint(7), Value{Type: TypeNumber, Number: 7}, false},
		{"NaN", math.NaN(), Value{}, true},
		{"Inf", math.Inf(1), Value{}, true},
		{"unsupported", []int{1}, Value{}, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := toValue(tt.input)
			if (err != nil) != tt.wantErr {
				t.Fatalf("toValue(%v) err = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
			if !tt.wantErr && got != tt.want {
				t.Errorf("toValue(%v) = %#v, want %#v", tt.input, got, tt.want)
			}
		})
	}
}
